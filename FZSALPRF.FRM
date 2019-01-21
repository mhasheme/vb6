VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRSALPERF 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Salary/Performance Review Report"
   ClientHeight    =   8955
   ClientLeft      =   570
   ClientTop       =   675
   ClientWidth     =   13050
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
   ScaleHeight     =   8955
   ScaleWidth      =   13050
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkIncIndSal 
      Caption         =   "Include Indirect Salary"
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
      Left            =   10320
      TabIndex        =   32
      Tag             =   "Check to Include Indirect Salary"
      Top             =   5760
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.CheckBox chkHideSalary 
      Caption         =   "Hide Salary"
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
      Left            =   8640
      TabIndex        =   31
      Top             =   6120
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.ComboBox comCountryOfEmp 
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
      Left            =   2230
      TabIndex        =   21
      Tag             =   "00-Country of Employment"
      Top             =   5205
      Width           =   1440
   End
   Begin VB.CheckBox chkPerPackaga 
      Caption         =   "Performance Package"
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
      Left            =   9750
      TabIndex        =   37
      Top             =   1560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "Show Reporting Authorities"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   5880
      TabIndex        =   29
      Top             =   6780
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "Show Position Description"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   5880
      TabIndex        =   28
      Top             =   6540
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "Show All Employees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   5880
      TabIndex        =   26
      Top             =   6060
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkDetails 
      Caption         =   "Show Comments"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   5880
      TabIndex        =   27
      Top             =   6300
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CheckBox chkInclHIS 
      Caption         =   "Include Performance History"
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
      Left            =   5880
      TabIndex        =   25
      Top             =   5820
      Visible         =   0   'False
      Width           =   2595
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2230
      MaxLength       =   4
      TabIndex        =   20
      Tag             =   "00-Employee Position Shift"
      Top             =   4860
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   2
      Left            =   5820
      TabIndex        =   19
      Tag             =   "00-Reporting Authority 3"
      Top             =   4530
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   18
      Tag             =   "00-Reporting Authority 2"
      Top             =   4530
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   17
      Tag             =   "00-Reporting Authority 1"
      Top             =   4530
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1920
      TabIndex        =   13
      Top             =   3210
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPC"
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   9480
      TabIndex        =   56
      Top             =   120
      Width           =   3375
      Begin Threed.SSOption optWhatDat 
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "Choose Between Salary or Performance Review Date"
         Top             =   600
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "    Next Performance Review Date"
         ForeColor       =   16711680
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
      Begin Threed.SSOption optWhatDat 
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "Choose Between Salary or Performance Review Date"
         Top             =   960
         Width           =   3015
         _Version        =   65536
         _ExtentX        =   5318
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "    Performance Review Date"
         ForeColor       =   16711680
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
      Begin Threed.SSOption optWhatDat 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Tag             =   "Choose Between Salary or Performance Review Date"
         Top             =   240
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   4895
         _ExtentY        =   344
         _StockProps     =   78
         Caption         =   "    Next Salary Review Date"
         ForeColor       =   16711680
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
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   55
      Top             =   8535
      Width           =   13050
      _Version        =   65536
      _ExtentX        =   23019
      _ExtentY        =   741
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   60
         Top             =   30
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
      Begin Crystal.CrystalReport vbxCrystal1 
         Left            =   3210
         Top             =   120
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
      Begin Crystal.CrystalReport vbxCrystal2 
         Left            =   3690
         Top             =   120
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
      Begin Crystal.CrystalReport vbxCrystal3 
         Left            =   4230
         Top             =   120
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
      Begin Crystal.CrystalReport vbxCrystal4 
         Left            =   4800
         Top             =   120
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
      Begin Crystal.CrystalReport vbxCrystal5 
         Left            =   5310
         Top             =   120
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
      Begin Crystal.CrystalReport vbxCrystal6 
         Left            =   5880
         Top             =   60
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
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
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
      Index           =   3
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Tag             =   "Final Sorting of Records"
      Top             =   7875
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
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
      Index           =   2
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Tag             =   "Third level of grouping records"
      Top             =   7530
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
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
      Index           =   1
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Tag             =   "Second level of grouping records"
      Top             =   7200
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
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
      Index           =   0
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Tag             =   "First Level of grouping records"
      Top             =   6870
      Width           =   2325
   End
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   1
      Left            =   3960
      TabIndex        =   23
      Tag             =   "Choose Between Detailed or Summary Report"
      Top             =   5820
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Summary"
      ForeColor       =   16711680
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   7
      Tag             =   "00-Enter Status Code"
      Top             =   1560
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Tag             =   "EDPT-Category"
      Top             =   1890
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
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Tag             =   "00-Enter Union Code"
      Top             =   1230
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   5
      Tag             =   "00-Enter Location Code"
      Top             =   900
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Tag             =   "00-Specific Department Desired"
      Top             =   570
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
      Left            =   1920
      TabIndex        =   3
      Tag             =   "00-Specific Division Desired"
      Top             =   240
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   10
      Tag             =   "00-Enter Position Group  Code"
      Top             =   2550
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1920
      TabIndex        =   15
      Tag             =   "00-Enter Administered By Code"
      Top             =   3870
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   16
      Tag             =   "00-Enter Section Code"
      Top             =   4200
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   14
      Tag             =   "00-Enter Region Code"
      Top             =   3540
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3750
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2880
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   2880
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1920
      TabIndex        =   9
      Tag             =   "10-Enter Employee Number"
      Top             =   2220
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   0
      Left            =   2220
      TabIndex        =   22
      TabStop         =   0   'False
      Tag             =   "Choose Between Detailed or Summary Report"
      Top             =   5820
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2302
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Detailed"
      ForeColor       =   16711680
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
   Begin Threed.SSCheck chkActing 
      Height          =   195
      Left            =   8640
      TabIndex        =   30
      Top             =   5820
      Visible         =   0   'False
      Width           =   1515
      _Version        =   65536
      _ExtentX        =   2672
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "Acting Position"
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
   Begin Threed.SSOption optGrouping 
      Height          =   195
      Index           =   2
      Left            =   2220
      TabIndex        =   24
      Tag             =   "Choose Between Detailed or Summary Report"
      Top             =   6120
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   344
      _StockProps     =   78
      Caption         =   "   Form"
      ForeColor       =   16711680
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
   Begin VB.Label lblCountry 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Country of Employment"
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
      Left            =   180
      TabIndex        =   62
      Top             =   5265
      Width           =   1620
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
      Left            =   180
      TabIndex        =   61
      Top             =   4905
      Visible         =   0   'False
      Width           =   645
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
      Left            =   180
      TabIndex        =   60
      Top             =   1935
      Width           =   630
   End
   Begin VB.Label lblRep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority:"
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
      Left            =   180
      TabIndex        =   59
      Top             =   4575
      Width           =   1395
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   180
      TabIndex        =   58
      Top             =   4245
      Width           =   540
   End
   Begin VB.Label lblPosGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
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
      Left            =   180
      TabIndex        =   57
      Top             =   2595
      Width           =   1035
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   180
      TabIndex        =   54
      Top             =   3585
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   180
      TabIndex        =   53
      Top             =   3915
      Width           =   1125
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   180
      TabIndex        =   52
      Top             =   945
      Width           =   615
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
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
      Left            =   210
      TabIndex        =   51
      Top             =   7935
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #3"
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
      Left            =   210
      TabIndex        =   50
      Top             =   7590
      Width           =   885
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #2"
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
      Left            =   210
      TabIndex        =   49
      Top             =   7260
      Width           =   885
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #1"
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
      Left            =   210
      TabIndex        =   48
      Top             =   6930
      Width           =   885
   End
   Begin VB.Label lblRepGrp 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Grouping"
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
      TabIndex        =   47
      Top             =   6555
      Width           =   1575
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Report"
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
      Index           =   8
      Left            =   180
      TabIndex        =   46
      Top             =   5820
      Width           =   1065
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
      TabIndex        =   45
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Performance Rating"
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
      Left            =   180
      TabIndex        =   44
      Top             =   3255
      Width           =   1410
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
      Left            =   180
      TabIndex        =   43
      Top             =   2925
      Width           =   1095
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
      Left            =   180
      TabIndex        =   42
      Top             =   2265
      Width           =   1290
   End
   Begin VB.Label lblCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Index           =   1
      Left            =   180
      TabIndex        =   41
      Top             =   1605
      Width           =   450
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   180
      TabIndex        =   40
      Top             =   1275
      Width           =   420
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
      Left            =   180
      TabIndex        =   39
      Top             =   615
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
      Left            =   180
      TabIndex        =   38
      Top             =   285
      Width           =   555
   End
End
Attribute VB_Name = "frmRSALPERF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public wdApp As Object 'As Word.Application
Public wrdDoc As Object 'Ticket #23876 Franks 06/0/2013
Dim fglbEmpList
Dim FNotMan, FYesMan

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Salary/" & lStr("Performance") & " Review Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    If glbCompSerial = "S/N - 2358W" Then 'City Of Belleville Ticket #23876 Franks 06/04/2013
        If optGrouping(2).Value Then
            Exit Sub
        End If
    End If
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    If glbLinamar And optGrouping(1) And chkPerPackaga Then
        Me.vbxCrystal1.Destination = 1
        Me.vbxCrystal6.Destination = 1
        Me.vbxCrystal2.Destination = 1
        Me.vbxCrystal3.Destination = 1
        Me.vbxCrystal4.Destination = 1
        Me.vbxCrystal5.Destination = 1
        
        Me.vbxCrystal5.Action = 1
        Me.vbxCrystal4.Action = 1
        Me.vbxCrystal3.Action = 1
        Me.vbxCrystal2.Action = 1
        If FNotMan Then Me.vbxCrystal1.Action = 1
        If FYesMan Then Me.vbxCrystal6.Action = 1
        
        Me.vbxCrystal1.Reset
        Me.vbxCrystal6.Reset
        Me.vbxCrystal2.Reset
        Me.vbxCrystal3.Reset
        Me.vbxCrystal4.Reset
        Me.vbxCrystal5.Reset
    End If
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"
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
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    If glbCompSerial = "S/N - 2358W" Then 'City Of Belleville Ticket #23876 Franks 06/04/2013
        If optGrouping(2).Value Then
            Exit Sub
        End If
    End If
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    If glbLinamar And optGrouping(1) And chkPerPackaga Then
        Me.vbxCrystal1.Destination = 0
        Me.vbxCrystal6.Destination = 0
        Me.vbxCrystal2.Destination = 0
        Me.vbxCrystal3.Destination = 0
        Me.vbxCrystal4.Destination = 0
        Me.vbxCrystal5.Destination = 0
        
        Me.vbxCrystal5.Action = 1
        Me.vbxCrystal4.Action = 1
        Me.vbxCrystal3.Action = 1
        Me.vbxCrystal2.Action = 1
        If FNotMan Then Me.vbxCrystal1.Action = 1
        If FYesMan Then Me.vbxCrystal6.Action = 1

        
        Me.vbxCrystal1.Reset
        Me.vbxCrystal6.Reset
        Me.vbxCrystal2.Reset
        Me.vbxCrystal3.Reset
        Me.vbxCrystal4.Reset
        Me.vbxCrystal5.Reset
    End If
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If

Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "SALPRF", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub chkHideSalary_Click()
'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
If glbCompSerial = "S/N - 2494W" Then
    If chkHideSalary Then
        chkIncIndSal.Visible = False
    Else
        chkIncIndSal.Visible = True
    End If
Else
    chkIncIndSal.Visible = False
End If
    
End Sub

Private Sub chkInclHIS_Click()
'If glbCompSerial = "S/N - 2228W" Then
    If chkInclHIS And (optWhatDat(1) Or optWhatDat(2)) Then
        chkDetails(0).Visible = True
        chkDetails(1).Visible = True
        chkDetails(2).Visible = True
        chkDetails(2).Value = 1
        chkDetails(3).Visible = True
    ElseIf chkInclHIS And optWhatDat(0) Then
        chkDetails(0).Visible = True
        chkDetails(1).Visible = True
        chkDetails(2).Visible = True
        chkDetails(2).Value = 1
        chkDetails(3).Visible = False
    Else
        chkDetails(0).Visible = False
        chkDetails(1).Visible = False
        chkDetails(2).Visible = False
        chkDetails(3).Visible = False
    End If

    'Release 8.0 - Ticket #22682: Hide Salary not custom reports
    If optGrouping(0) Or optGrouping(2) Then
        chkHideSalary.Visible = False
    Else
        If (optWhatDat(1) Or optWhatDat(2)) And chkInclHIS Then
            chkHideSalary.Visible = False
        Else
            chkHideSalary.Visible = True
        End If
    End If

    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If chkInclHIS Then
            chkIncIndSal.Visible = False
        Else
            chkIncIndSal.Visible = True
        End If
    Else
        chkIncIndSal.Visible = False
    End If

'End If
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
Dim vPosGroup As String
    
    If Not glbSyndesis Then
        vPosGroup = "Position Group"
    Else
        vPosGroup = "Position Grade"
    End If
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("G/L")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Administered By")
    comGroup(0).AddItem vPosGroup '"Position Group" 'Frank 4/19/2000
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    'If glbLinamar Then comGroup(0).AddItem "Home Line"
    comGroup(0).AddItem "Employee Name"
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem IIf(lStr("PShift") = "PShift", "Shift", lStr("PShift"))
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem lStr("Rept. Authority 1")  'Ticket #22682: Release 8.0
    comGroup(0).AddItem "(none)"
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem lStr("Performance") & " Rating"
    comGroup(1).AddItem "(none)"
    
    comGroup(2).AddItem lStr("Performance") & " Rating"
    comGroup(2).AddItem "(none)"
    comGroup(3).AddItem "Next Salary/" & lStr("Performance") & " Review Date"
    
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    comGroup(3).ListIndex = 0
    comGroup(3).Enabled = False
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 1 Then strCd$ = IIf(glbMulti, "HR_JOB_HISTORY.JH_ORG", "HREMP.ED_ORG")
    If intIdx% = 2 Then strCd$ = "HREMP.ED_EMP"
    If intIdx% = 3 Then strCd$ = "HRJOB.JB_GRPCD"
    If intIdx% = 4 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 6 Then strCd$ = "HREMP.ED_SECTION"
    If intIdx% = 7 Then strCd$ = "HR_PERFORM_HISTORY.PH_PCODE"
    
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
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
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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
    EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
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

Private Sub Cri_RepAuth()
Dim TempCri As String
Dim EECri As String, LocCri As String
Dim I, xTemp As Boolean
Dim xField

    xTemp = False
    EECri = ""
    xField = ""

    If optWhatDat(0) Then
        xField = "{HR_JOB_HISTORY.JH_REPTAU"
    Else
        xField = "{HR_PERFORM_HISTORY.PH_REPTAU"
    End If
    
    If Len(Trim(elpRept(0).Text)) > 0 Then
        EECri = EECri & xField & "} = " & getEmpnbr(Trim(elpRept(0).Text)) & " "
        xTemp = True
    End If
    If Len(Trim(elpRept(1).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "and " & xField & "2} = " & getEmpnbr(Trim(elpRept(1).Text)) & " "
        Else
            EECri = EECri & xField & "2} = " & getEmpnbr(Trim(elpRept(1).Text)) & " "
        End If
        xTemp = True
    End If
    If Len(Trim(elpRept(2).Text)) > 0 Then
        If xTemp Then
            EECri = EECri & "and " & xField & "3} = " & getEmpnbr(Trim(elpRept(2).Text)) & " "
        Else
            EECri = EECri & xField & "3} = " & getEmpnbr(Trim(elpRept(2).Text)) & " "
        End If
        xTemp = True
    End If
                
    
    If Len(EECri) > 0 Then
        If Len(glbstrSelCri) > 0 Then
          glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
          glbstrSelCri = EECri
        End If
    End If

End Sub

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim X%, strCriNam$

If optWhatDat(0).Value = True Then
    strCriNam$ = "HR_SALARY_HISTORY.SH_NEXTDAT"
Else
    If optWhatDat(1).Value = True Then
        strCriNam$ = "HR_PERFORM_HISTORY.PH_PNEXT"
    Else 'optWhatDat(2) Ticket #13188 Frank 06/12/07
        strCriNam$ = "HR_PERFORM_HISTORY.PH_PREVIEW"
    End If
End If


If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({" & strCriNam$ & "} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For X% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({" & strCriNam$ & "} "
        If X% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next X%


Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub
If glbMulti Then
    EECri = "{HR_JOB_HISTORY.JH_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
Else
    EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_ActingPosition()
Dim TempCri As String
Dim EECri As String, LocCri As String
Dim I, xTemp As Boolean
    xTemp = False
    EECri = ""
    If chkActing.Value = 0 Or optGrouping(0) Then
        Exit Sub
    End If

    EECri = EECri & "{HR_JOB_HISTORY.JH_POSITION_CONTROL} = 'YES' "
    xTemp = True

    If Len(EECri) > 0 Then
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
    End If

End Sub

Private Function Cri_SetAll()
Dim X%, XFMT, LocstrSelCri
Dim strSelCri14

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
strSelCri14 = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_CountryOfEmployment 'Ticket #13621 Frank 09/07/2007
Call Cri_RepAuth
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 7
    Call Cri_Code(X%)
Next X%

If chkActing.Visible = True Then
    Call Cri_ActingPosition
End If

If optGrouping(1) And glbLinamar Then 'And (chkPerPackaga And chkPerPackaga.Visible) Then
    Call AttWrk
    strSelCri14 = glbstrSelCri   'Ticket #15316
End If

Call Cri_FTDates

If glbCompSerial = "S/N - 2358W" Then 'City Of Belleville Ticket #23876 Franks 06/04/2013
    If optGrouping(2).Value Then
        Call BellevilleAppraisalWrk
        Call Generate_Word_Forms_Belleville
        Exit Function
    End If
End If

'Ticket #28635 - Add View Own security
'If (optWhatDat(1) Or optWhatDat(2)) Then
    If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_Performance_ViewOwn Then
        'Do not show user's Performance records based on the Employee # associated to the User.
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND {HR_PERFORM_HISTORY.PH_EMPNBR} <> " & glbUserEmpNo
        Else
            glbstrSelCri = glbstrSelCri & " {HR_PERFORM_HISTORY.PH_EMPNBR} <> " & glbUserEmpNo
        End If
    End If
'End If

' report name
If optGrouping(0).Value = True Then
    Call setRptLabel(Me, 1)
    If glbMulti Then
        If glbCompSerial = "S/N - 2368W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "ACLsalpdt.rpt"
        Else
            'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
            Me.vbxCrystal.Formulas(13) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
            
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmsalpdt.rpt"
        End If
    Else
        'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
        Me.vbxCrystal.Formulas(13) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
    
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsalpdt.rpt"
    End If
    
    If Not glbLinamar Then
        Me.vbxCrystal.Formulas(10) = "lblPerformance = '" & lStr("Performance") & "'"
    End If
    
    'x% = Cri_Sorts()
    LocstrSelCri = ""
    If glbNoNONE And glbNoEXEC Then     'Hemu -EXE
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG } <> 'EXEC'))"
        Else
            LocstrSelCri = "(isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG } <> 'NONE' AND {HREMP.ED_ORG } <> 'EXEC'))"
        End If
    ElseIf glbNoNONE Then
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
        Else
            LocstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'NONE')"
        End If
    ElseIf glbNoEXEC Then   'Hemu -EXE
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"
        Else
            LocstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'EXEC')"
        End If
    Else
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = glbstrSelCri
        End If
    End If
    If glbWFC Then
        If gSec_WFC_Band_Security Then
            LocstrSelCri = LocstrSelCri & " AND ({HR_SALARY_HISTORY.SH_BAND}='A' OR {HR_SALARY_HISTORY.SH_BAND}='B' OR {HR_SALARY_HISTORY.SH_BAND}='C' OR {HR_SALARY_HISTORY.SH_BAND}='D' OR {HR_SALARY_HISTORY.SH_BAND}='E')"
        End If
    End If
    
    If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
        'Limit the user has access to these employees salary information
        If Len(glbNoAccessGrp) <> 0 Then
            LocstrSelCri = LocstrSelCri & " AND (" & glbNoAccessGrp & ") "
            LocstrSelCri = Replace(Replace(LocstrSelCri, "HREMP.ED_", "ED_"), "ED_", "HREMP.ED_")
            LocstrSelCri = Replace(LocstrSelCri, "(HREMP.", "({HREMP.")
            LocstrSelCri = Replace(LocstrSelCri, " HREMP.", " {HREMP.")
            LocstrSelCri = Replace(LocstrSelCri, " =", "} =")
        End If
    End If
    
    Me.vbxCrystal.SelectionFormula = LocstrSelCri
    'If Len(glbstrSelCri) >= 0 Then Me.vbxCrystal.SelectionFormula = glbstrSelCri

Else
    If glbLinamar Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rlslprfs.rpt"
    ElseIf glbMulti Then
        If glbCompSerial = "S/N - 2368W" Then
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "ACLslprfs.rpt"
        Else
            If optWhatDat(0) Then
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmslprfs.rpt"
            Else
                Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmslprfsP.rpt"
            End If
            
            'Release 8.0 - Ticket #22682: Hide Salary
            If chkHideSalary.Visible Then
                Me.vbxCrystal.Formulas(12) = "hideSalary=" & (chkHideSalary <> 0)
            End If
            
            'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
            Me.vbxCrystal.Formulas(13) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
        End If
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzslprfs.rpt"
        
        'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
        Me.vbxCrystal.Formulas(13) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
        
        'Release 8.0 - Ticket #22682: Hide Salary
        If (optWhatDat(1) Or optWhatDat(2)) And chkInclHIS Then
            'don't set the value for Hide Salary
        Else
            If chkHideSalary.Visible Then
                Me.vbxCrystal.Formulas(12) = "hideSalary=" & (chkHideSalary <> 0)
            End If
        End If
    End If
    If Not glbLinamar Then
        If optWhatDat(1) Or optWhatDat(2) Then
            If chkInclHIS Then
                'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
                Me.vbxCrystal.Formulas(13) = ""
            
                If chkDetails(1) Then
                    If glbMulti Then
                        If glbCompSerial = "S/N - 2368W" Then
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "ACLprfhs1.rpt"
                        Else
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmprfhs1.rpt"
                        End If
                    Else
                        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzprfhs1.rpt"
                    End If
                Else
                    If glbMulti Then
                        If glbCompSerial = "S/N - 2368W" Then
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "ACLprfhis.rpt"
                        Else
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmprfhis.rpt"
                        End If
                    Else
                        If chkActing Then
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzprfhisAct.rpt"
                        Else
                            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzprfhis.rpt"
                        End If
                    End If
                End If
            End If
        ElseIf optWhatDat(0) Then
            If chkInclHIS Then
                If chkDetails(1) Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "slrhs1.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "slrhis.rpt"
                End If
                
                'Release 8.0 - Ticket #22682: Hide Salary
                If chkHideSalary.Visible Then
                    Me.vbxCrystal.Formulas(12) = "hideSalary=" & (chkHideSalary <> 0)
                End If
                
                'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
                Me.vbxCrystal.Formulas(13) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "
            End If
        End If
    End If
    
    If Not glbLinamar Then
        Me.vbxCrystal.Formulas(10) = "lblPerformance = '" & lStr("Performance") & "'"
    End If
    
    X% = Cri_Sorts()
    
    LocstrSelCri = ""
    If glbNoNONE And glbNoEXEC Then     'Hemu -EXE
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
        Else
            LocstrSelCri = "(isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG } <> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
        End If
    ElseIf glbNoNONE Then
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
        Else
            LocstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'NONE')"
        End If
    ElseIf glbNoEXEC Then     'Hemu -EXE
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"
        Else
            LocstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'EXEC')"
        End If
    Else
        If Len(glbstrSelCri) >= 0 Then
            LocstrSelCri = glbstrSelCri
        End If
    End If
    
    If glbWFC Then
        If gSec_WFC_Band_Security Then
            If chkInclHIS = 0 Then
                LocstrSelCri = LocstrSelCri & " AND ({HR_SALARY_HISTORY.SH_BAND}='A' OR {HR_SALARY_HISTORY.SH_BAND}='B' OR {HR_SALARY_HISTORY.SH_BAND}='C' OR {HR_SALARY_HISTORY.SH_BAND}='D' OR {HR_SALARY_HISTORY.SH_BAND}='E')"
            End If
        End If
    End If
    
    If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
        'Limit the user has access to these employees salary information
        If Len(glbNoAccessGrp) <> 0 Then
            If Len(LocstrSelCri) > 0 Then
                LocstrSelCri = LocstrSelCri & " AND (" & glbNoAccessGrp & ") "
            Else
                LocstrSelCri = LocstrSelCri & " (" & glbNoAccessGrp & ") "
            End If
            LocstrSelCri = Replace(Replace(LocstrSelCri, "HREMP.ED_", "ED_"), "ED_", "HREMP.ED_")
            LocstrSelCri = Replace(LocstrSelCri, "(HREMP.", "({HREMP.")
            LocstrSelCri = Replace(LocstrSelCri, " HREMP.", " {HREMP.")
            LocstrSelCri = Replace(LocstrSelCri, " =", "} =")
        End If
    End If
    
    Me.vbxCrystal.SelectionFormula = LocstrSelCri
        
End If
Me.vbxCrystal.Connect = RptODBC_SQL

If glbLinamar And optGrouping(1) Then
    
    Me.vbxCrystal1.ReportFileName = glbIHRREPORTS & "RLPrfEvl.rpt"
    Me.vbxCrystal6.ReportFileName = glbIHRREPORTS & "RLMrfEvl.rpt"
    Me.vbxCrystal2.ReportFileName = glbIHRREPORTS & "RLEPWrkS.rpt"
    Me.vbxCrystal3.ReportFileName = glbIHRREPORTS & "RLPayChg.rpt"
    Me.vbxCrystal4.ReportFileName = glbIHRREPORTS & "RLATTSHT.rpt"
    Me.vbxCrystal5.ReportFileName = glbIHRREPORTS & "RLEDSEM.rpt"
    
    Me.vbxCrystal1.SelectionFormula = glbstrSelCri & " AND NOT ({HRJOB.JB_GRPCD}='MAN' OR {HRJOB.JB_GRPCD}='EXEC' OR {HRJOB.JB_GRPCD}='SUP')"
    Me.vbxCrystal6.SelectionFormula = glbstrSelCri & " AND ({HRJOB.JB_GRPCD}='MAN' OR {HRJOB.JB_GRPCD}='EXEC' OR {HRJOB.JB_GRPCD}='SUP')"
    Me.vbxCrystal2.SelectionFormula = glbstrSelCri
    Me.vbxCrystal3.SelectionFormula = glbstrSelCri
    
    'Ticket #15316 and then Ticket #15466
    Me.vbxCrystal4.SelectionFormula = glbstrSelCri & " AND {LN_ATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    'Me.vbxCrystal4.SelectionFormula = strSelCri14 & " AND {LN_ATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    
    Me.vbxCrystal5.SelectionFormula = glbstrSelCri
  
    Me.vbxCrystal1.WindowTitle = "Employee " & lStr("Performance") & " Evaluation"
    Me.vbxCrystal6.WindowTitle = "Managerial Employee " & lStr("Performance") & " Evaluation"
    Me.vbxCrystal2.WindowTitle = "Employee Preparation Work Sheet 1 & 2"
    Me.vbxCrystal3.WindowTitle = "Payroll Change Notice"
    Me.vbxCrystal4.WindowTitle = "Annual Employee Attendance Sheet"
    Me.vbxCrystal5.WindowTitle = "Education and Seminars Information"
    

    Me.vbxCrystal1.Connect = RptODBC_SQL
    Me.vbxCrystal6.Connect = RptODBC_SQL
    Me.vbxCrystal2.Connect = RptODBC_SQL
    Me.vbxCrystal3.Connect = RptODBC_SQL
    Me.vbxCrystal4.Connect = RptODBC_SQL
    Me.vbxCrystal5.Connect = RptODBC_SQL
End If

Me.vbxCrystal.WindowTitle = "Salary and " & lStr("Performance") & " Review Report"

Cri_SetAll = True


Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%, SavGrp1, SavGrp2, SavGrp3, SavGrp4

Cri_Sorts = 0

' first set primary grouping
grpField$ = getEGroup(comGroup(0).Text)
If comGroup(0).Text = "(none)" Then grpField$ = "{HRPARCO.PC_NAME}"
If comGroup(0) <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descName1 = " & grpField$ 'added by RAUBREY 5/15/97
    Me.vbxCrystal.Formulas(1) = dscGroup$ 'added by RAUBREY 5/15/97
Else
    Me.vbxCrystal.Formulas(0) = "descGroup1= ''"
    Me.vbxCrystal.Formulas(1) = "descName1 = ''"
End If

grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$
If glbLinamar Then
    Me.vbxCrystal1.GroupCondition(0) = grpCond$
    Me.vbxCrystal6.GroupCondition(0) = grpCond$
    Me.vbxCrystal2.GroupCondition(0) = grpCond$
End If

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{tblPerform.TB_DESC}"
    Case 2: grpField$ = "{@EFullName}"
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$
If glbLinamar Then
    Me.vbxCrystal1.GroupCondition(1) = grpCond$
    Me.vbxCrystal6.GroupCondition(1) = grpCond$
    Me.vbxCrystal2.GroupCondition(1) = grpCond$
End If

GrpIdx% = comGroup(2).ListIndex
Select Case GrpIdx%
     Case 0: grpField$ = "{tblPerform.TB_DESC}"
     Case 1
        If optWhatDat(0) Then
            grpField$ = "{HR_SALARY_HISTORY.SH_NEXTDAT}"
        Else
            grpField$ = "{HR_PERFORM_HISTORY.PH_PNEXT}"
        End If
End Select

grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(2) = grpCond$
If glbLinamar Then
    Me.vbxCrystal1.GroupCondition(2) = grpCond$
    Me.vbxCrystal6.GroupCondition(2) = grpCond$
    Me.vbxCrystal2.GroupCondition(2) = grpCond$
End If

If Not glbLinamar Then
'Hemu - Simcoe
'If glbCompSerial = "S/N - 2228W" Then
    If optWhatDat(1) Or optWhatDat(2) Then
        If chkInclHIS Then
            Me.vbxCrystal.Formulas(2) = "showComm=" & (chkDetails(0) <> 0)
            Me.vbxCrystal.Formulas(3) = "showPos=" & (chkDetails(2) <> 0)
            Me.vbxCrystal.Formulas(4) = "showRpt=" & (chkDetails(3) <> 0)
        End If
    ElseIf optWhatDat(0) Then
        If chkInclHIS Then
            Me.vbxCrystal.Formulas(2) = "showComm=" & (chkDetails(0) <> 0)
            Me.vbxCrystal.Formulas(3) = "showPos=" & (chkDetails(2) <> 0)
        End If
    End If
'End If
'Hemu - Simcoe
End If

Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim X%, I

CriCheck = False

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

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For X% = 0 To 7
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

For X% = 0 To 1
    If Len(dlpDateRange(X%).Text) > 0 Then
        If Not IsDate(dlpDateRange(X%).Text) Then
            MsgBox "Not a valid date"
            dlpDateRange(X%).Text = ""
            dlpDateRange(X%).SetFocus
            Exit Function
        End If
    End If
Next X%
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If

For I = 0 To 2
    If elpRept(I).Caption = "Enter Valid Employee #" Then
        MsgBox "If Reporting Authority Entered - they must exist"
        elpRept(I).SetFocus
        Exit Function
    End If
Next

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Private Sub Form_Activate()

Call SET_UP_MODE

'Hamilton CAS - Ticket #13398
If glbCompSerial = "S/N - 2257W" Then
    If chkActing.Visible = True Then
        chkActing.Caption = "Red Circled"
    End If
End If

If glbCompSerial = "S/N - 2358W" Then 'City Of Belleville Ticket #23876 Franks 06/04/2013
    optGrouping(2).Visible = True
End If

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

'Ticket #16228 - Oxford
If glbCompSerial = "S/N - 2259W" Then
    glbMulti = True
End If

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
    lblShift.Caption = lStr("PShift")
    If lblShift.Caption = "PShift" Then lblShift.Caption = "Shift"
End If

Call comGrpLoad
Call setRptCaption(Me)

'Ticket #24122 - Label Master
chkActing.Caption = lStr(chkActing.Caption)

If glbLinamar Then clpCode(4).MaxLength = 8
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6
If glbMulti Then lblUnion.ForeColor = &HC000C0: lblPT.ForeColor = &HC000C0

If optGrouping(1) Then
    If optWhatDat(0) Then
        chkInclHIS.Caption = "Include Salary History"
    Else
        chkInclHIS.Caption = "Include " & lStr("Performance") & " History"
    End If
    chkInclHIS.Visible = True
    
    chkActing.Visible = True
        
Else
    chkActing.Visible = False
End If

Call INI_Controls(Me)

Me.Caption = lStr(Me.Caption)
optWhatDat(1).Caption = lStr(optWhatDat(1).Caption)
optWhatDat(2).Caption = lStr(optWhatDat(2).Caption)
lblBCode.Caption = lStr(lblBCode.Caption)
'chkInclHIS.Caption = lStr(chkInclHIS.Caption)

If glbCompSerial = "S/N - 2368W" Then
    optWhatDat(1).Value = True
    optWhatDat(0).Value = False
End If

'Hamilton CAS - Ticket #13398
If glbCompSerial = "S/N - 2257W" Then
    If chkActing.Visible = True Then
        chkActing.Caption = "Red Circled"
    End If
End If

If Not glbCompSerial = "S/N - 2358W" Then 'City Of Belleville Ticket #23876 Franks 06/04/2013
    optGrouping(2).Visible = False
End If

'Release 8.0 - Ticket #22682: Hide Salary not custom reports
If glbLinamar Or glbCompSerial = "S/N - 2358W" Then
    chkHideSalary.Visible = False
End If

'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
If glbCompSerial = "S/N - 2494W" Then
    chkIncIndSal.Visible = True
Else
    chkIncIndSal.Value = False
    chkIncIndSal.Visible = False
End If

Call addCountryItems 'Frank 09/07/2007 Ticket #13621

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
If glbLinamar Then
    gdbAdoIhr001.Execute "DELETE FROM LN_ATTWRK WHERE AD_WRKEMP='" & glbUserID & "'"
End If
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub

Private Sub optGrouping_Click(Index As Integer, Value As Integer)
'If glbCompSerial = "S/N - 2228W" Then
    If optGrouping(1) Then
        chkInclHIS.Visible = True
        chkInclHIS.Value = False
        chkDetails(0).Value = False
        chkDetails(1).Value = False
        chkDetails(2).Value = False
        chkDetails(3).Value = False
        
        chkActing.Visible = True
        
        'Hamilton CAS - Ticket #13398
        If glbCompSerial = "S/N - 2257W" Then
            If chkActing.Visible = True Then
                chkActing.Caption = "Red Circled"
            End If
        End If
    Else
        chkInclHIS.Visible = False
        chkDetails(0).Visible = False
        chkDetails(1).Visible = False
        chkDetails(2).Visible = False
        chkDetails(3).Visible = False
        chkActing.Visible = False
    End If
'End If
    If glbLinamar Then
        If optGrouping(1) Then
            If optWhatDat(1) Or optWhatDat(2) Then
                chkPerPackaga.Visible = True
            Else
                chkPerPackaga.Visible = False
            End If
        Else
            chkPerPackaga.Visible = False
        End If
    End If
    
    'Release 8.0 - Ticket #22682: Hide Salary not custom reports
    If optGrouping(0) Or optGrouping(2) Then
        chkHideSalary.Visible = False
    Else
        If (optWhatDat(1) Or optWhatDat(2)) And chkInclHIS Then
            chkHideSalary.Visible = False
        Else
            chkHideSalary.Visible = True
        End If
    End If
End Sub

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optWhatDat_Click(Index As Integer, Value As Integer)
'If glbCompSerial = "S/N - 2228W" Then
    If (optWhatDat(1) Or optWhatDat(0) Or optWhatDat(2)) And optGrouping(1) Then
        If optWhatDat(0) Then
            chkInclHIS.Caption = "Include Salary History"
        Else
            chkInclHIS.Caption = "Include " & lStr("Performance") & " History"
        End If
        
        chkInclHIS.Visible = True
        chkInclHIS.Value = False
        chkDetails(0).Value = False
        chkDetails(1).Value = False
        chkDetails(2).Value = False
        chkDetails(3).Value = False
        If glbLinamar Then
            If optWhatDat(1) Or optWhatDat(2) Then
                chkPerPackaga.Visible = True
            Else
                chkPerPackaga.Visible = False
            End If
        End If
    Else
        chkInclHIS.Visible = False
        chkDetails(0).Visible = False
        chkDetails(1).Visible = False
        chkDetails(2).Visible = False
        chkDetails(3).Visible = False
        If glbLinamar Then
            chkPerPackaga.Visible = False
        End If
    End If
    
    'Release 8.0 - Ticket #22682: Hide Salary not custom reports
    If optGrouping(0) Or optGrouping(2) Then
        chkHideSalary.Visible = False
    Else
        If (optWhatDat(1) Or optWhatDat(2)) And chkInclHIS Then
            chkHideSalary.Visible = False
        Else
            chkHideSalary.Visible = True
        End If
    End If
    
'End If
End Sub

Private Sub optWhatDat_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub AttWrk()
Dim WSQLQ, strSFormat$
Dim DATE1, DATE2

DATE2 = Date

If IsDate(dlpDateRange(0).Text) Then
    DATE2 = dlpDateRange(0).Text
End If
DATE1 = DateAdd("d", -1, DateAdd("yyyy", -1, DATE2))

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 50
MDIMain.panHelp(1).Caption = " Please Wait"
    
gdbAdoIhr001.CommandTimeout = 600
WSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "'", "''"), "{", ""), "}", ""), "[", "("), "]", ")")
'jaddy changed for reducing the length to less than 500 because the parameters only allow 500 chat
WSQLQ = Replace(WSQLQ, "1=1 and 1=1", "1=1")
WSQLQ = Replace(WSQLQ, "1=1 and ", "")
WSQLQ = Replace(WSQLQ, " and 1=1", "")
'end
gdbAdoIhr001.Execute "QRY_LN_ATTWRK '" & WSQLQ & "','" & glbUserID & "','" & Format(DATE1, "mmm dd,yyyy") & "','" & Format(DATE2, "mmm dd,yyyy") & "'"
gdbAdoIhr001.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 90
Call AboutManager
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
If DATE1 <> "" And DATE2 <> "" Then
    strSFormat$ = "As of " & DATE1 & " through " & DATE2
    Me.vbxCrystal4.Formulas(1) = "Daterange = '" & strSFormat$ & "'"
Else
    strSFormat$ = "No date entered"
    Me.vbxCrystal4.Formulas(1) = "Daterange = '" & strSFormat$ & "'"
End If
End Sub

Private Sub AboutManager()
Dim rsMAN As New ADODB.Recordset, SQLQ
Dim WSQLQ, ISQLQ

WSQLQ = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")
ISQLQ = " SELECT ED_EMPNBR FROM HREMP "
ISQLQ = ISQLQ & " INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR "
ISQLQ = ISQLQ & " INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB=HRJOB.JB_CODE "

If optWhatDat(0).Value = True Then
    ISQLQ = ISQLQ & " INNER JOIN HR_SALARY_HISTORY ON HREMP.ED_EMPNBR=HR_SALARY_HISTORY.SH_EMPNBR  "
    ISQLQ = ISQLQ & " WHERE SH_CURRENT<>0"
    If IsDate(dlpDateRange(0).Text) Then ISQLQ = ISQLQ & " AND SH_NEXTDAT  >= " & Date_SQL(dlpDateRange(0))
    If IsDate(dlpDateRange(1).Text) Then ISQLQ = ISQLQ & " AND SH_NEXTDAT  <= " & Date_SQL(dlpDateRange(1))
Else
    ISQLQ = ISQLQ & " INNER JOIN HR_PERFORM_HISTORY ON HREMP.ED_EMPNBR=HR_PERFORM_HISTORY.PH_EMPNBR  "
    ISQLQ = ISQLQ & " WHERE PH_CURRENT<>0"
    If IsDate(dlpDateRange(0).Text) Then ISQLQ = ISQLQ & " AND PH_PNEXT  >= " & Date_SQL(dlpDateRange(0))
    If IsDate(dlpDateRange(1).Text) Then ISQLQ = ISQLQ & " AND PH_PNEXT  <= " & Date_SQL(dlpDateRange(1))
End If
ISQLQ = ISQLQ & " AND JH_CURRENT<>0 AND " & WSQLQ

SQLQ = ISQLQ & " AND JB_GRPCD IN ('MAN','SUP','EXEC') "
rsMAN.Open SQLQ, gdbAdoIhr001, adOpenStatic
FYesMan = Not rsMAN.EOF
rsMAN.Close
SQLQ = ISQLQ & " AND JB_GRPCD NOT IN ('MAN','SUP','EXEC') "
rsMAN.Open SQLQ, gdbAdoIhr001, adOpenStatic
FNotMan = Not rsMAN.EOF
rsMAN.Close

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

Private Sub Cri_CountryOfEmployment()
Dim CountryCri As String

If Len(comCountryOfEmp.Text) > 0 Then
    CountryCri = "({HREMP.ED_WORKCOUNTRY} = '" & comCountryOfEmp.Text & "')"
End If

If Len(CountryCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = CountryCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & CountryCri
    End If
    glbiOneWhere = True
End If
End Sub

Private Sub addCountryItems()
Dim ctylist, X

ctylist = CountryList
X = 1
Do While X > 0
    X = InStr(ctylist, "&")
    If X > 0 Then
        comCountryOfEmp.AddItem Left(ctylist, X - 1)
        ctylist = Mid(ctylist, X + 1)
    Else
        comCountryOfEmp.AddItem ctylist
    End If
Loop

End Sub

Private Function CountryList() As String
Dim xCountryList As String, ctyFile
xCountryList = ""
ctyFile = glbIHRREPORTS & "CountryList.MTF"

On Error GoTo ErrorHandler

If File(ctyFile) Then
    Open ctyFile For Input As #1
    Input #1, xCountryList
    Close #1
End If

ResumeHere:
'If InStr(xCountryList, BasicCountry) = 0 Then
'    xCountryList = BasicCountry
'End If
If InStr(xCountryList, comCountryOfEmp) = 0 And comCountryOfEmp <> "" Then
    xCountryList = xCountryList & "&" & comCountryOfEmp
    comCountryOfEmp.AddItem comCountryOfEmp
End If
Open ctyFile For Output As #1
Print #1, xCountryList
Close #1
CountryList = xCountryList
Exit Function

ErrorHandler:
If Err.Number = 62 Then
    ' Corrupted CountryList.MTF, kill it and regenerate
    Close #1
    MsgBox "Found corrupt CountryList.MTF.  info:HR will re-create this file.", vbInformation + vbOKOnly, "Corrupted Country List"
    Kill ctyFile
    Resume ResumeHere
Else
    'MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number & " in CountryList"
    Resume Next
End If
End Function

Private Function Generate_Word_Forms_Belleville() 'City Of Belleville Ticket #23876 Franks 06/04/2013
    Dim rsEmp As New ADODB.Recordset
    Dim rsWRK As New ADODB.Recordset
    Dim rsEPos As New ADODB.Recordset
    Dim rsESal As New ADODB.Recordset
    Dim rsEPerf As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim SQLQ As String
    Dim xEmpNo As String
    Dim xlsFileTmpl1, xlsFileTmpl2, xlsFileTmpl3, xlsFileTmpl4, xlsFileTmpl5, xlsFileTmplM
    Dim xlsFileMat As String
    Dim xlsFileMain As String
    Dim xSaveFileName As String
    Dim xRow, xRecNum As Integer
    Dim strReptAuth1 As String
    Dim I As Long
    Dim totNum As Long
    Dim empName, EmpDocName
    Dim xFName
    Dim locPath, xLocPath
    Dim xJob, xJobDesc
    Dim xPerNextReview
    Dim xStep4, xStep5, xTemp
    Dim xStep4Amt, xStep5Amt 'Ticket #24093 Franks 07/18/2013
    Dim xSaveToPath
    Dim strRName$
    
    On Error GoTo Err_Word_Forms
    
    Screen.MousePointer = HOURGLASS

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."
    
    SQLQ = "DELETE FROM PAYROLL_TRANSFER_RPT WHERE [TABLE] = 'BellevilleApp'"
    gdbAdoIhr001.Execute SQLQ
    Call Pause(1)
    
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' ORDER BY TT_ID"
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsWRK.EOF Then
        MsgBox "No record found."
        Exit Function
    End If
    I = 0
    If Not rsWRK.EOF Then totNum = rsWRK.RecordCount
    
    'Begin ------------------------------------------------------------------
    locPath = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\")
    'Get the Word Template
    xlsFileTmpl1 = locPath & "CBellevilleAppraisal.dot"
    If Dir(xlsFileTmpl1) = "" Then
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(0).FloodPercent = 0
        MDIMain.panHelp(0).Caption = ""
        Screen.MousePointer = DEFAULT
        MsgBox "There is no " & xlsFileTmpl1
        Exit Function
    End If
    
    xSaveToPath = "Q:\HR_GENERAL\PERFORMANCE EVALUATIONS"
    ''xSaveToPath = "C:\INFOHR\Database Files 7.9\MS SQL SERVER\PERFORMANCE EVALUATIONS"
    If Dir(xSaveToPath, vbDirectory) = "" Then
        xSaveToPath = locPath
    Else
        xSaveToPath = xSaveToPath & "\"
    End If
    
    'Create new hidden instance of Word.
    Set wdApp = CreateObject("Word.Application")

    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
        I = I + 1
        DoEvents
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & rsWRK("TT_EMPNBR") & " "
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmp.EOF Then
        
            empName = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
            EmpDocName = rsEmp("ED_FNAME") & " " & rsEmp("ED_SURNAME")
            xFName = "CBellevilleAppraisal(" & empName & ").doc"
            'xlsFileMat = locPath & "" & xFName
            xlsFileMat = xSaveToPath & "" & xFName
            
            'delete this file if it exists
            If (Dir(xlsFileMat)) <> "" Then
                Kill xlsFileMat
            End If
            xSaveFileName = xlsFileMat
                
            xJob = ""
            xJobDesc = ""
            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            If rsEPos.State <> 0 Then rsEPos.Close
            rsEPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEPos.EOF Then
                xJob = rsEPos("JH_JOB")
                xJobDesc = GetJobDesc(xJob)
            End If
            rsEPos.Close
            
            xPerNextReview = ""
            SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE NOT PH_CURRENT = 0 AND PH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            If rsEPerf.State <> 0 Then rsEPerf.Close
            rsEPerf.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEPerf.EOF Then
                If Not IsNull(rsEPerf("PH_PNEXT")) Then
                    xPerNextReview = rsEPerf("PH_PNEXT")
                End If
            End If
            rsEPerf.Close
                
            xStep4 = ""
            xStep5 = ""
            xStep4Amt = ""
            xStep5Amt = ""
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            If rsESal.State <> 0 Then rsESal.Close
            rsESal.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsESal.EOF Then
                If Not IsNull(rsESal("SH_GRADE")) Then
                    xStep4 = rsESal("SH_GRADE")
                    xStep4 = Val(xStep4)
                    xStep4Amt = Format(rsESal("SH_SALARY"), "$0.00")
                    xStep5 = xStep4 + 1
                    'check if there is step 5
                    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & rsESal("SH_JOB") & "' "
                    If rsJOB.State <> 0 Then rsJOB.Close
                    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If rsJOB.EOF Then
                        xStep5 = ""
                    Else
                        If xStep5 > 11 Then
                            xStep5 = ""
                        Else
                            xTemp = "JB_S" & Trim(Str(xStep5))
                            If IsNull(rsJOB(xTemp)) Then
                                xStep5 = "No Change"
                                xStep5Amt = "No Change"
                            ElseIf rsJOB(xTemp) = 0 Then
                                xStep5 = "No Change"
                                xStep5Amt = "No Change"
                            Else
                                If rsJOB(xTemp) > 0 Then
                                    xStep5Amt = Format(rsJOB(xTemp), "$0.00")
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            rsESal.Close
            'If Len(xStep4) > 0 Then xStep4 = "PAY RATE, STEP " & xStep4
            'If Len(xStep5) > 0 Then
            '    If IsNumeric(xStep5) Then xStep5 = "PAY RATE, STEP " & xStep5
            'End If
            If Len(xStep4) > 0 Then xStep4Amt = xStep4Amt & ", STEP " & xStep4
            If Len(xStep5) > 0 And Not (xStep5 = "No Change") Then
                If IsNumeric(xStep5) Then xStep5Amt = xStep5Amt & ", STEP" & xStep5
            End If
            
            With wdApp
                'Set Word object as the template
                Set wrdDoc = .Documents.Add(xlsFileTmpl1, False)

                'Show this instance of Word.
                '.Visible = True

                'Make the word doc Active
                .Documents(wrdDoc).Activate

                'Update the bookmark fields in the Word template with database values
                .ActiveDocument.FormFields("txtName").Result = EmpDocName
                .ActiveDocument.FormFields("txtDept").Result = getDeptDesc(rsEmp("ED_DEPTNO"))
                .ActiveDocument.FormFields("txtPos").Result = xJobDesc
                .ActiveDocument.FormFields("txtRevDate").Result = xPerNextReview
                .ActiveDocument.FormFields("txtStep4").Result = xStep4Amt ' xStep4 '"Step 4" '
                .ActiveDocument.FormFields("txtStep5").Result = xStep5Amt 'xStep5
               
                'Save the template as the Word Document - with the filename generated above
                wrdDoc.SaveAs xlsFileMat

                .ActiveDocument.Close
                'Set wrdDoc = Nothing

            End With
                
            If Len(xlsFileMat) > 0 Then
                'add this employee and file name in the report
                Call ExpLocTRANSFERRPT(rsEmp("ED_EMPNBR"), empName, "Word File Name:", xFName, "BellevilleApp", 1, "") 'xDETAILS2)  '"")
            End If
            
            ''If Len(xlsFileMat) > 0 Then
            ''    Call Pause(1)
            ''    'launch the file
            ''    'Shell "Start " & GetShortName(xlsFileMat)
            ''    If Not LanchXlsW98(xlsFileMat) Then
            ''        Shell "cmd /c " & GetShortName(xlsFileMat)
            ''    End If
            ''End If
    
        End If
        rsWRK.MoveNext
    Loop
    rsWRK.Close
    
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Set wrdDoc = Nothing
    wdApp.NormalTemplate.Saved = True
    wdApp.Quit
    Set wdApp = Nothing
        
    strRName$ = glbIHRREPORTS & "rzPayTransImp.rpt"
    Me.vbxCrystal.ReportFileName = strRName$
    Me.vbxCrystal.Formulas(10) = "rptTitle='Merge Appraisal Document Created List'"
    Me.vbxCrystal.Formulas(11) = "NoteOfTitle='This folder keeps the following Word files: " & xSaveToPath & "' "
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.SelectionFormula = "{PAYROLL_TRANSFER_RPT.TABLE}='BellevilleApp'"
    Me.vbxCrystal.WindowTitle = "Merge Appraisal Document Created List"
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
        
        
    Screen.MousePointer = DEFAULT
    
Exit Function

Err_Word_Forms:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    If Err = 53 Then
        Resume Next
    End If
    If Err = 1004 Then
        Resume Next
    End If
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close all reports and try again.", vbExclamation + vbOKOnly, "Error generating report"
        GoTo Close_All
    End If
    If Err = 70 Then
        MsgBox Err.Description & Chr(10) & "Please close all reports and try again.", vbExclamation + vbOKOnly, "Error generating report"
        GoTo Close_All
    End If
    
    'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hourly Benefit Statement", "Hourly Benefit Statement", "Generate Word Form")
    Resume Next
    'wrdDoc.Close wdDoNotSaveChanges
    
Close_All:
    'wdApp.NormalTemplate.Saved = True
''    wdApp.Quit
''    Set wdApp = Nothing
''
''    wdAppMain.NormalTemplate.Saved = True
''    wdAppMain.Quit
''    Set wdAppMain = Nothing
    
    'Generate_Word_Forms_Hourly = -1
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT

End Function

Private Sub BellevilleAppraisalWrk()
Dim WSQLQ, strSFormat$
Dim DATE1, DATE2
Dim rsEmp As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsEPos As New ADODB.Recordset
Dim rsESal As New ADODB.Recordset
Dim rsEPerf As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ
Dim I As Long
Dim totNum As Long
Dim xSkip As Boolean
On Error GoTo Mod_Err
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = " Please Wait"
        
    'Delete Temp tables
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    WSQLQ = "WHERE " & glbSeleDeptUn & " "
    If Len(clpDiv.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "') "
    If Len(clpDept.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "') "
    If Len(clpCode(0).Text) > 0 Then WSQLQ = WSQLQ & "AND ED_LOC='" & clpCode(0).Text & "' "
    If Len(clpCode(1).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "') "
    If Len(clpCode(2).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMP IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
    If Len(clpPT.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "') "
    If Len(elpEEID.Text) > 0 Then WSQLQ = WSQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    If Len(clpCode(4).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_REGION IN ('" & Replace(clpCode(4).Text, ",", "','") & "') "
    If Len(clpCode(5).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_ADMINBY IN ('" & Replace(clpCode(5).Text, ",", "','") & "') "
    If Len(clpCode(6).Text) > 0 Then WSQLQ = WSQLQ & " AND ED_SECTION IN ('" & Replace(clpCode(6).Text, ",", "','") & "') "
    If Len(txtShift.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_SHIFT='" & txtShift.Text & "' "
    If Len(comCountryOfEmp.Text) > 0 Then WSQLQ = WSQLQ & "AND ED_WORKCOUNTRY='" & comCountryOfEmp.Text & "' "
    SQLQ = "SELECT * FROM HREMP "
    SQLQ = SQLQ & WSQLQ
    SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
    If rsEmp.State <> 0 Then rsEmp.Close
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then totNum = rsEmp.RecordCount
    I = 0
    Do While Not rsEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
        I = I + 1
        DoEvents
        xSkip = False
        'check position
        If Len(clpCode(3).Text) > 0 Or Len(elpRept(0).Text) > 0 Or Len(elpRept(1).Text) > 0 Or Len(elpRept(2).Text) > 0 Then '
            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
            If rsEPos.State <> 0 Then rsEPos.Close
            rsEPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsEPos.EOF Then
                xSkip = True
            End If
            If Not rsEPos.EOF Then
                If Len(clpCode(3).Text) > 0 Then 'Position Group
                    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & rsEPos("JH_JOB") & "' "
                    If rsJOB.State <> 0 Then rsJOB.Close
                    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not rsJOB.EOF Then
                        If Not IsNull(rsJOB("JB_GRPCD")) Then
                            If Not clpCode(3).Text = rsJOB("JB_GRPCD") Then
                                xSkip = True
                            End If
                        End If
                    End If
                    rsJOB.Close
                End If
                If Len(elpRept(0).Text) > 0 Then
                    If Not IsNull(rsEPos("JH_REPTAU")) Then
                        If Not elpRept(0).Text = rsEPos("JH_REPTAU") Then
                            xSkip = True
                        End If
                    Else
                        xSkip = True
                    End If
                End If
                If Len(elpRept(1).Text) > 0 Then
                    If Not IsNull(rsEPos("JH_REPTAU2")) Then
                        If Not elpRept(1).Text = rsEPos("JH_REPTAU2") Then
                            xSkip = True
                        End If
                    Else
                        xSkip = True
                    End If
                End If
                If Len(elpRept(2).Text) > 0 Then
                    If Not IsNull(rsEPos("JH_REPTAU3")) Then
                        If Not elpRept(2).Text = rsEPos("JH_REPTAU3") Then
                            xSkip = True
                        End If
                    Else
                        xSkip = True
                    End If
                End If
            End If
            rsEPos.Close
            
        End If
        
        'salary
        If optWhatDat(0).Value Then
            If Len(dlpDateRange(0).Text) > 0 Or Len(dlpDateRange(1).Text) > 0 Then
                SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
                If rsESal.State <> 0 Then rsESal.Close
                rsESal.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If rsESal.EOF Then
                    xSkip = True
                End If
                If Not rsESal.EOF Then
                    If IsDate(dlpDateRange(0).Text) Then
                        If Not IsNull(rsESal("SH_NEXTDAT")) Then
                            If Not CVDate(rsESal("SH_NEXTDAT")) >= CVDate(dlpDateRange(0).Text) Then
                                xSkip = True
                            End If
                        Else
                            xSkip = True
                        End If
                    End If
                    If IsDate(dlpDateRange(1).Text) Then
                        If Not IsNull(rsESal("SH_NEXTDAT")) Then
                            If Not CVDate(rsESal("SH_NEXTDAT")) <= CVDate(dlpDateRange(1).Text) Then
                                xSkip = True
                            End If
                        Else
                            xSkip = True
                        End If
                    End If
                End If
                rsESal.Close
            End If
        End If
        
        'performance
        If optWhatDat(1).Value Or optWhatDat(2).Value Then
            If Len(dlpDateRange(0).Text) > 0 Or Len(dlpDateRange(1).Text) > 0 Or Len(clpCode(7).Text) > 0 Then
                SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE NOT PH_CURRENT = 0 AND PH_EMPNBR = " & rsEmp("ED_EMPNBR") & " "
                If rsEPerf.State <> 0 Then rsEPerf.Close
                rsEPerf.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If rsEPerf.EOF Then
                    xSkip = True
                End If
                If Not rsEPerf.EOF Then
                    If optWhatDat(1).Value Then '    Next Performance Review Date
                        If IsDate(dlpDateRange(0).Text) Then
                            If Not IsNull(rsEPerf("PH_PNEXT")) Then
                                If Not CVDate(rsEPerf("PH_PNEXT")) >= CVDate(dlpDateRange(0).Text) Then
                                    xSkip = True
                                End If
                            Else
                                xSkip = True
                            End If
                        End If
                        If IsDate(dlpDateRange(1).Text) Then
                            If Not IsNull(rsEPerf("PH_PNEXT")) Then
                                If Not CVDate(rsEPerf("PH_PNEXT")) <= CVDate(dlpDateRange(1).Text) Then
                                    xSkip = True
                                End If
                            Else
                                xSkip = True
                            End If
                        End If
                    End If
                    If optWhatDat(2).Value Then '    Performance Review Date
                        If IsDate(dlpDateRange(0).Text) Then
                            If Not IsNull(rsEPerf("PH_PREVIEW")) Then
                                If Not CVDate(rsEPerf("PH_PREVIEW")) >= CVDate(dlpDateRange(0).Text) Then
                                    xSkip = True
                                End If
                            Else
                                xSkip = True
                            End If
                        End If
                        If IsDate(dlpDateRange(1).Text) Then
                            If Not IsNull(rsEPerf("PH_PREVIEW")) Then
                                If Not CVDate(rsEPerf("PH_PREVIEW")) <= CVDate(dlpDateRange(1).Text) Then
                                    xSkip = True
                                End If
                            Else
                                xSkip = True
                            End If
                        End If
                    End If
                    If Len(clpCode(7).Text) > 0 Then 'Performance Rating
                            If Not IsNull(rsEPerf("PH_PCODE")) Then
                                If Not rsEPerf("PH_PCODE") = clpCode(7).Text Then
                                    xSkip = True
                                End If
                            Else
                                xSkip = True
                            End If
                    End If
                End If
                rsEPerf.Close
            End If
        End If
        If Not xSkip Then
            'add to wrk table
            rsWRK.AddNew
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
        End If
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    MDIMain.panHelp(0).FloodPercent = 100
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(2).Caption = ""
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = " "
    Exit Sub
Mod_Err:
    MsgBox Err.Description
    'Resume Next
End Sub

Private Function getDeptDesc(xCode) 'Ticket #23876 Franks 06/10/2013
Dim rsDEPT As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRDEPT WHERE DF_NBR = '" & xCode & "' "
        rsDEPT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDEPT.EOF Then
            xRetVal = rsDEPT("DF_NAME")
        End If
        rsDEPT.Close
    End If
    getDeptDesc = xRetVal
End Function
Private Function GetJobDesc(xCode) 'Ticket #23876 Franks 06/10/2013
Dim rsJOB As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
        rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsJOB.EOF Then
            xRetVal = rsJOB("JB_DESCR")
        End If
        rsJOB.Close
    End If
    GetJobDesc = xRetVal
End Function

'Ticket #23876 Franks 06/13/2013
Private Sub ExpLocTRANSFERRPT(xEmpNo, xEmpName, xDETAILS1, xMessage, xTabFunc, xImp, xDETAILS2, Optional xYear)
Dim rsRPT As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM PAYROLL_TRANSFER_RPT WHERE [TABLE] ='" & xTabFunc & "' "
    SQLQ = SQLQ & "AND EMPNBR = " & xEmpNo & " "
    'SQLQ = SQLQ & "AND DETAILS1 = '" & xSIN & "' "
    'If Not IsMissing(xYear) Then
    '    SQLQ = SQLQ & "AND TABLE_ID = " & xYear & " "
    'End If
    rsRPT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsRPT.EOF Then
        rsRPT.AddNew
    End If
    rsRPT("TABLE") = xTabFunc '"Manulife_Export"
    If Not IsMissing(xYear) Then
        rsRPT("TABLE_ID") = xYear
    End If
    rsRPT("MESSAGE") = xMessage
    rsRPT("IMPORTED") = xImp
    rsRPT("EMPNBR") = xEmpNo
    rsRPT("EMPNAME") = xEmpName
    'rsRPT("EDSECTION") = xDiv
    rsRPT("DETAILS1") = xDETAILS1
    rsRPT("DETAILS2") = xDETAILS2
    'If Not IsMissing(xDETAILS3) Then
    '    rsRPT("DETAILS3") = xDETAILS3
    'End If
    rsRPT("LDATE") = Date
    rsRPT("LTIME") = "Not Show"
    rsRPT.Update
    rsRPT.Close
End Sub

