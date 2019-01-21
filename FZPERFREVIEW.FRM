VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRPRFREVIW
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Salary/Performance Review Report"
   ClientHeight    =   8130
   ClientLeft      =   570
   ClientTop       =   675
   ClientWidth     =   10785
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
   ScaleHeight     =   8130
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
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
      TabIndex        =   54
      Top             =   6420
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
      TabIndex        =   53
      Top             =   6180
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
      TabIndex        =   52
      Top             =   5700
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
      TabIndex        =   51
      Top             =   5940
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
      Height          =   285
      Left            =   5880
      TabIndex        =   50
      Top             =   5220
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2235
      MaxLength       =   4
      TabIndex        =   17
      Tag             =   "00-Employee Position Shift"
      Top             =   4860
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Index           =   2
      Left            =   5820
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      TabIndex        =   10
      Top             =   3210
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPC"
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   5760
      TabIndex        =   44
      Top             =   240
      Width           =   3375
      Begin Threed.SSOption optWhatDat 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   20
         Tag             =   "Choose Between Salary or Performance Review Date"
         Top             =   240
         Width           =   2775
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "    Next Salary Review Date"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   31.73
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin Threed.SSOption optWhatDat 
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   21
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
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   43
      Top             =   7710
      Width           =   10785
      _Version        =   65536
      _ExtentX        =   19024
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
      TabIndex        =   25
      Tag             =   "Final Sorting of Records"
      Top             =   6510
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
      TabIndex        =   24
      Tag             =   "Third level of grouping records"
      Top             =   6195
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
      TabIndex        =   23
      Tag             =   "Second level of grouping records"
      Top             =   5880
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
      TabIndex        =   22
      Tag             =   "First Level of grouping records"
      Top             =   5550
      Width           =   2325
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   1
      Left            =   3960
      TabIndex        =   19
      Tag             =   "Choose Between Detailed or Summary Report"
      Top             =   5190
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "   Summary"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
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
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1560
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   1890
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1230
      Width           =   4005
      _ExtentX        =   7064
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
      TabIndex        =   2
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
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   570
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   240
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   7
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
      TabIndex        =   12
      Tag             =   "00-Enter Administered By Code"
      Top             =   3870
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   13
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
      TabIndex        =   11
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2220
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   6955
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   0
      Left            =   2220
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "Choose Between Detailed or Summary Report"
      Top             =   5190
      Width           =   1305
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "   Detailed"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      TabIndex        =   49
      Top             =   4890
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
      TabIndex        =   48
      Top             =   1890
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
      TabIndex        =   47
      Top             =   4560
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
      TabIndex        =   46
      Top             =   4230
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
      TabIndex        =   45
      Top             =   2550
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
      TabIndex        =   42
      Top             =   3570
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
      TabIndex        =   41
      Top             =   3900
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
      TabIndex        =   40
      Top             =   900
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
      TabIndex        =   39
      Top             =   6600
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
      TabIndex        =   38
      Top             =   6300
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
      TabIndex        =   37
      Top             =   6000
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
      TabIndex        =   36
      Top             =   5715
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
      TabIndex        =   35
      Top             =   5480
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
      TabIndex        =   34
      Top             =   5220
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
      TabIndex        =   33
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
      TabIndex        =   32
      Top             =   3240
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
      TabIndex        =   31
      Top             =   2910
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
      TabIndex        =   30
      Top             =   2220
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
      TabIndex        =   29
      Top             =   1560
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
      TabIndex        =   28
      Top             =   1230
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
      TabIndex        =   27
      Top             =   570
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
      TabIndex        =   26
      Top             =   240
      Width           =   555
   End
End
Attribute VB_Name = "frmRPRFREVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEmpList
Dim FNotMan, FYesMan
Private Sub cmdClose_Click()
Unload Me
End Sub



Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Salary/" & lStr("Performance") & " Review Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    If glbLinamar And optGrouping(1) Then
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
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    If glbLinamar And optGrouping(1) Then
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

Private Sub chkInclHIS_Click()
'If glbCompSerial = "S/N - 2228W" Then
    If chkInclHIS And optWhatDat(1) Then
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
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
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
    DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
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
Dim x%, strCriNam$

If optWhatDat(0).Value = True Then
    strCriNam$ = "HR_SALARY_HISTORY.SH_NEXTDAT"
Else
    strCriNam$ = "HR_PERFORM_HISTORY.PH_PNEXT"
End If



If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({" & strCriNam$ & "} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = Month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = Month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({" & strCriNam$ & "} "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = Month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        GoTo Cri_FTDatst
    End If
Next x%



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
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub
If glbMulti Then
    EECri = "{HR_JOB_HISTORY.JH_PT}= '" & clpPT.Text & "'"
Else
    EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"
End If

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True


End Sub

Private Function Cri_SetAll()
Dim x%, XFMT, locstrSelCri

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_RepAuth
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 7
    Call Cri_Code(x%)
Next x%
If optGrouping(1) Then If glbLinamar Then Call AttWrk
Call Cri_FTDates
' report name
If optGrouping(0).Value = True Then
    Call setRptLabel(Me, 1)
    If glbMulti Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmsalpdt.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzsalpdt.rpt"
    End If
    
    If Not glbLinamar Then
        Me.vbxCrystal.Formulas(10) = "lblPerformance = '" & lStr("Performance") & "'"
    End If
    
    'x% = Cri_Sorts()
        locstrSelCri = ""
        If glbNoNONE And glbNoEXEC Then     'Hemu -EXE
            If Len(glbstrSelCri) >= 0 Then
                locstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG } <> 'EXEC'))"
            Else
                locstrSelCri = "(isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG } <> 'NONE' AND {HREMP.ED_ORG } <> 'EXEC'))"
            End If
        ElseIf glbNoNONE Then
            If Len(glbstrSelCri) >= 0 Then
                locstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
            Else
                locstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'NONE')"
            End If
        ElseIf glbNoEXEC Then   'Hemu -EXE
            If Len(glbstrSelCri) >= 0 Then
                locstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"
            Else
                locstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'EXEC')"
            End If
        Else
            If Len(glbstrSelCri) >= 0 Then
                locstrSelCri = glbstrSelCri
            End If
        End If
        If glbWFC Then
            If gSec_WFC_Band_Security Then
                locstrSelCri = locstrSelCri & " AND ({HR_SALARY_HISTORY.SH_BAND}='A' OR {HR_SALARY_HISTORY.SH_BAND}='B' OR {HR_SALARY_HISTORY.SH_BAND}='C' OR {HR_SALARY_HISTORY.SH_BAND}='D' OR {HR_SALARY_HISTORY.SH_BAND}='E')"
            End If
        End If
        Me.vbxCrystal.SelectionFormula = locstrSelCri
    'If Len(glbstrSelCri) >= 0 Then Me.vbxCrystal.SelectionFormula = glbstrSelCri

Else
    If glbLinamar Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rlslprfs.rpt"
    ElseIf glbMulti Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rmslprfs.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzslprfs.rpt"
    End If
    If Not glbLinamar Then
        If optWhatDat(1) Then
            If chkInclHIS Then
                If chkDetails(1) Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "prfhs1.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "prfhis.rpt"
                End If
            End If
        ElseIf optWhatDat(0) Then
            If chkInclHIS Then
                If chkDetails(1) Then
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "slrhs1.rpt"
                Else
                    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & IIf(glbMulti, "rm", "rz") & "slrhis.rpt"
                End If
            End If
        End If
    End If
    
    If Not glbLinamar Then
        Me.vbxCrystal.Formulas(10) = "lblPerformance = '" & lStr("Performance") & "'"
    End If
    
    x% = Cri_Sorts()
    locstrSelCri = ""
    If glbNoNONE And glbNoEXEC Then     'Hemu -EXE
        If Len(glbstrSelCri) >= 0 Then
            locstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
        Else
            locstrSelCri = "(isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG } <> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
        End If
    ElseIf glbNoNONE Then
        If Len(glbstrSelCri) >= 0 Then
            locstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
        Else
            locstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'NONE')"
        End If
    ElseIf glbNoEXEC Then     'Hemu -EXE
        If Len(glbstrSelCri) >= 0 Then
            locstrSelCri = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"
        Else
            locstrSelCri = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'EXEC')"
        End If
    Else
        If Len(glbstrSelCri) >= 0 Then
            locstrSelCri = glbstrSelCri
        End If
    End If
    
    If glbWFC Then
        If gSec_WFC_Band_Security Then
            If chkInclHIS = 0 Then
                locstrSelCri = locstrSelCri & " AND ({HR_SALARY_HISTORY.SH_BAND}='A' OR {HR_SALARY_HISTORY.SH_BAND}='B' OR {HR_SALARY_HISTORY.SH_BAND}='C' OR {HR_SALARY_HISTORY.SH_BAND}='D' OR {HR_SALARY_HISTORY.SH_BAND}='E')"
            End If
        End If
    End If
    Me.vbxCrystal.SelectionFormula = locstrSelCri
        
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
    Me.vbxCrystal4.SelectionFormula = glbstrSelCri & " AND {LN_ATTWRK.AD_WRKEMP}='" & glbUserID & "'"
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
Dim EECri As String, OneSet%, x%

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
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
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
    If optWhatDat(1) Then
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
Dim x%, I

CriCheck = False

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If


For x% = 0 To 7
    If Not clpCode(x).ListChecker Then Exit Function

Next x%


For x% = 0 To 1
 If Len(dlpDateRange(x%).Text) > 0 Then
    If Not IsDate(dlpDateRange(x%).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(x%).Text = ""
        dlpDateRange(x%).SetFocus
        Exit Function
    End If
 End If
Next x%
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
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call comGrpLoad
Call setRptCaption(Me)
If glbLinamar Then clpCode(4).MaxLength = 8
If glbCompSerial = "S/N - 2381W" Then clpCode(0).MaxLength = 6
If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6
If glbMulti Then lblUnion.ForeColor = &HC000C0: lblPT.ForeColor = &HC000C0

If optGrouping(1) Then
    If optWhatDat(0) Then
        chkInclHIS.Caption = "Include Salary History"
    Else
        chkInclHIS.Caption = "Include " & lStr("Performance") & " History"
    End If
    chkInclHIS.Visible = True
End If

Call INI_Controls(Me)

Me.Caption = lStr(Me.Caption)
optWhatDat(1).Caption = lStr(optWhatDat(1).Caption)
lblBCode.Caption = lStr(lblBCode.Caption)
'chkInclHIS.Caption = lStr(chkInclHIS.Caption)

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
    Else
        chkInclHIS.Visible = False
        chkDetails(0).Visible = False
        chkDetails(1).Visible = False
        chkDetails(2).Visible = False
        chkDetails(3).Visible = False
    End If
'End If
End Sub

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub optWhatDat_Click(Index As Integer, Value As Integer)
'If glbCompSerial = "S/N - 2228W" Then
    If (optWhatDat(1) Or optWhatDat(0)) And optGrouping(1) Then
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
    Else
        chkInclHIS.Visible = False
        chkDetails(0).Visible = False
        chkDetails(1).Visible = False
        chkDetails(2).Visible = False
        chkDetails(3).Visible = False
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
    
gdbAdoIhr001.CommandTimeout = 300
WSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "'", "''"), "{", ""), "}", ""), "[", "("), "]", ")")
'jaddy changed for reducing the length to less than 500 because the parameters only allow 500 chat
WSQLQ = Replace(WSQLQ, "1=1 and 1=1", "1=1")
WSQLQ = Replace(WSQLQ, "1=1 and ", "")
WSQLQ = Replace(WSQLQ, " and 1=1", "")
'end
gdbAdoIhr001.Execute "QRY_LN_ATTWRK '" & WSQLQ & "','" & glbUserID & "','" & Format(DATE1, "mmm dd,yyyy") & "','" & Format(DATE2, "mmm dd,yyyy") & "'"
gdbAdoIhr001.CommandTimeout = 300
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

