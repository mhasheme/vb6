VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUTERM 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employee Termination"
   ClientHeight    =   9345
   ClientLeft      =   315
   ClientTop       =   780
   ClientWidth     =   12645
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
   ScaleHeight     =   9345
   ScaleWidth      =   12645
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   8
      Tag             =   "10-Enter Employee Number"
      Top             =   3060
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7375
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Threed.SSPanel Panel3D1 
      Height          =   3090
      Left            =   0
      TabIndex        =   19
      Top             =   4200
      Width           =   10275
      _Version        =   65536
      _ExtentX        =   18124
      _ExtentY        =   5450
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
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Left            =   2100
         TabIndex        =   10
         Tag             =   "41-Date Terminated"
         Top             =   630
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   2100
         TabIndex        =   11
         Tag             =   "41-Termination Code - Code "
         Top             =   990
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TERM"
      End
      Begin Threed.SSCheck chkSum 
         Height          =   225
         Left            =   6120
         TabIndex        =   14
         Tag             =   "Click to Select Summarize Attendance Records    "
         Top             =   1440
         Width           =   2865
         _Version        =   65536
         _ExtentX        =   5054
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "Summarize Attendance Records      "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin Threed.SSCheck chkRehire 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Tag             =   "Click to Select Rehire"
         Top             =   1320
         Width           =   2505
         _Version        =   65536
         _ExtentX        =   4419
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Rehire                                       "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Value           =   -1  'True
      End
      Begin VB.TextBox txtComments 
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
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   15
         Tag             =   "00-Comments - free form"
         Top             =   1920
         Width           =   8895
      End
      Begin VB.CommandButton cmdTerminate 
         Appearance      =   0  'Flat
         Caption         =   "Terminate  Employees"
         Height          =   330
         Left            =   6480
         TabIndex        =   16
         Tag             =   "Terminate the Employee Selected"
         Top             =   840
         Width           =   2220
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   7
         Left            =   2100
         TabIndex        =   12
         Tag             =   "00-Termination Cause"
         Top             =   240
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TECA"
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Cause"
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
         Index           =   7
         Left            =   150
         TabIndex        =   44
         Top             =   330
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.Label lblTermData 
         Caption         =   "Termination Data"
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
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label lblRehire 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Yes"
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
         Height          =   255
         Left            =   3480
         TabIndex        =   21
         Top             =   1380
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments"
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
         Left            =   180
         TabIndex        =   20
         Top             =   1710
         Width           =   870
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Tag             =   "41-Date Terminated"
         Top             =   690
         Width           =   1470
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   18
         Top             =   1020
         Width           =   1710
      End
   End
   Begin Threed.SSPanel panTermRpts 
      Height          =   2835
      Left            =   5640
      TabIndex        =   31
      Top             =   360
      Width           =   4155
      _Version        =   65536
      _ExtentX        =   7329
      _ExtentY        =   5001
      _StockProps     =   15
      Caption         =   "Termination Reports"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Alignment       =   0
      Begin VB.CommandButton cmdPrintSelected 
         Appearance      =   0  'Flat
         Caption         =   "Print Selected Reports"
         Height          =   330
         Left            =   900
         TabIndex        =   32
         Tag             =   "Print the reports marked with an 'x'"
         Top             =   2040
         Width           =   2220
      End
      Begin Threed.SSCheck chkTermRpts 
         Height          =   225
         Index           =   5
         Left            =   330
         TabIndex        =   33
         Top             =   1450
         Width           =   1845
         _Version        =   65536
         _ExtentX        =   3254
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "   Employee Comments"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   225
         Index           =   4
         Left            =   330
         TabIndex        =   34
         Tag             =   "Click to Select Follow-Ups"
         Top             =   1240
         Width           =   2085
         _Version        =   65536
         _ExtentX        =   3678
         _ExtentY        =   397
         _StockProps     =   78
         Caption         =   "   Follow-Ups"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   255
         Index           =   3
         Left            =   330
         TabIndex        =   35
         Tag             =   "Click to Select Entitlements with Compensatory Time, Hourly Entitlements"
         Top             =   830
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Entitlements with Compensatory "
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   255
         Index           =   2
         Left            =   330
         TabIndex        =   36
         Tag             =   "Click to Select Employee Profile"
         Top             =   590
         Width           =   1545
         _Version        =   65536
         _ExtentX        =   2725
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Employee Profile"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   255
         Index           =   0
         Left            =   330
         TabIndex        =   37
         Tag             =   "Click to Select Attendance History"
         Top             =   350
         Width           =   1995
         _Version        =   65536
         _ExtentX        =   3519
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Attendance History"
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
      Begin Threed.SSCheck chkTermRpts 
         Height          =   225
         Index           =   1
         Left            =   330
         TabIndex        =   39
         Tag             =   "Click to select Compensatory Time"
         Top             =   1670
         Visible         =   0   'False
         Width           =   1965
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "   Compensatory Time "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   27
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "Time, Hourly Entitlements"
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
         Left            =   720
         TabIndex        =   42
         Top             =   1045
         Width           =   2295
      End
      Begin VB.Label lblRptsPrinted 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reports Printed for this Employee"
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
         Left            =   960
         TabIndex        =   38
         Top             =   2400
         Visible         =   0   'False
         Width           =   2340
      End
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   750
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   420
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   1410
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   5
      Tag             =   "00-Enter Location Code"
      Top             =   2070
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Tag             =   "EDPT-Category"
      Top             =   1740
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   6
      Tag             =   "00-Enter Region Code"
      Top             =   2400
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Enter Administered By Code"
      Top             =   2730
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   600
      Top             =   6840
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
      Index           =   6
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   3405
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
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
      Left            =   60
      TabIndex        =   43
      Top             =   3480
      Width           =   540
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
      Left            =   60
      TabIndex        =   41
      Top             =   1800
      Width           =   630
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
      Left            =   60
      TabIndex        =   40
      Top             =   3150
      Width           =   1290
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Administrated by"
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
      TabIndex        =   29
      Top             =   2835
      Width           =   1155
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   60
      TabIndex        =   28
      Top             =   2475
      Width           =   510
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
      Left            =   60
      TabIndex        =   27
      Top             =   480
      Width           =   555
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
      Left            =   60
      TabIndex        =   26
      Top             =   810
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
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
      TabIndex        =   25
      Top             =   1140
      Width           =   840
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
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
      TabIndex        =   24
      Top             =   1470
      Width           =   1350
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
      Left            =   60
      TabIndex        =   23
      Top             =   2130
      Width           =   615
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
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmUTERM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lblEEID, I
Dim Title$, EID&, TermDate$
Dim HisSQL, HisSQL1
Dim MailBody

Private Function AUDITTERM()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String, XSNAME As String, XFNAME As String, XEMPTYPE As String
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITTERM = False

'Hemu - Added this because on Mass Termination also we have to transfer to Vadim - confirmed with Jerry.
glbChgTermReason = clpCode(1)
glbChgTermDate = dlpTermDate
Call TermPayrollEmp(dlpTermDate, glbLEE_ID)


rsTB.Open "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME, ED_EMPTYPE,ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
If Not rsTB.EOF Then
    'xPT = rsTB("ED_PT")
    'xDiv = rsTB("ED_DIV")
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
    XSNAME = rsTB("ED_SURNAME")
    XFNAME = rsTB("ED_FNAME")
    XEMPTYPE = IIf(IsNull(rsTB("ED_EMPTYPE")), "", rsTB("ED_EMPTYPE")) 'George Apr 13,2006
Else
    xPT = ""
    xDiv = ""
    XSNAME = ""
    XFNAME = ""
    XEMPTYPE = ""
End If

strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, "
strFields = strFields & "AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_EMPTYPE, AU_SURNAME, "
strFields = strFields & "AU_FNAME, AU_DOT, AU_TREAS, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM"
rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP"
rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL"
rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_EMPTYPE") = XEMPTYPE
rsTA("AU_SURNAME") = XSNAME
rsTA("AU_FNAME") = XFNAME
rsTA("AU_DOT") = dlpTermDate.Text
rsTA("AU_TREAS") = clpCode(1).Text
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = "T"
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If
rsTA.Update

AUDITTERM = True
GoTo AUDITTERM_CLOSE
'Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js
AUDITTERM_CLOSE:
    rsTA.Close
    rsTB.Close
End Function

Private Sub chkRehire_Click(Value As Integer)

If chkRehire.Value = True Then
    lblRehire.Caption = "Yes"
Else
    lblRehire.Caption = "No"
End If

End Sub

Private Sub chkRehire_GotFocus()
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Sub chkSum_Click(Value As Integer)

If chkSum.Value = False Then
    glbchkSum = False
Else
    glbchkSum = True
End If

End Sub

Private Function chkTerms(RptTo As Integer)
Dim dd As Integer

chkTerms = False

If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("Invalid Division Code")
     clpDiv.SetFocus
    Exit Function
End If
If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "Invalid Department Code"
     clpDept.SetFocus
    Exit Function
End If
If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If
If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
    MsgBox lStr("Invalid Union Code")
    clpCode(0).SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) > 0 And clpCode(2).Caption = "Unassigned" Then
    MsgBox "Invalid Employment Status"
    clpCode(2).SetFocus
    Exit Function
End If
If Len(clpCode(3).Text) > 0 And clpCode(3).Caption = "Unassigned" Then
    MsgBox lStr("Invalid Location Code")
    clpCode(3).SetFocus
    Exit Function
End If
If Len(clpCode(4).Text) > 0 And clpCode(4).Caption = "Unassigned" Then
    MsgBox lStr("Invalid Region Code")
    clpCode(4).SetFocus
    Exit Function
End If
If Len(clpCode(5).Text) > 0 And clpCode(5).Caption = "Unassigned" Then
    MsgBox "Invalid Admin By Code"
    clpCode(5).SetFocus
    Exit Function
End If
If Len(clpCode(6).Text) > 0 And clpCode(6).Caption = "Unassigned" Then
    MsgBox "Invalid Section Code"
    clpCode(6).SetFocus
    Exit Function
End If

If RptTo <> 1 Then
    If Len(dlpTermDate.Text) < 1 Then
        MsgBox "Termination Date is a required field"
        dlpTermDate.SetFocus
        Exit Function
    End If
    
    If Not IsDate(dlpTermDate.Text) Then
        MsgBox "Termination Date is not a valid date."
        dlpTermDate.SetFocus
        Exit Function
    End If
    
    If Len(clpCode(1).Text) < 1 Then
            MsgBox "Termination Reason is a required field"
            clpCode(1).SetFocus
            Exit Function
    Else
        If Len(clpCode(1).Text) > 1 And clpCode(1).Caption = "Unassigned" Then
            MsgBox "Invalid Termination Reason "
            clpCode(1).SetFocus
            Exit Function
        End If
    End If

End If
If Not elpEEID.ListChecker Then
    Exit Function
End If
Dim HasSelectionCriteria As Boolean
HasSelectionCriteria = False
If Len(clpDept.Text) > 0 Then HasSelectionCriteria = True
If Len(clpDiv.Text) > 0 Then HasSelectionCriteria = True
If Len(clpCode(0).Text) > 0 Then HasSelectionCriteria = True
If Len(clpPT.Text) > 0 Then HasSelectionCriteria = True
If Len(clpCode(2).Text) > 0 Then HasSelectionCriteria = True
If Len(clpCode(3).Text) > 0 Then HasSelectionCriteria = True
If Len(clpCode(4).Text) > 0 Then HasSelectionCriteria = True
If Len(clpCode(5).Text) > 0 Then HasSelectionCriteria = True
If Len(clpCode(6).Text) > 0 Then HasSelectionCriteria = True
If Len(elpEEID.Text) > 0 Then HasSelectionCriteria = True
If Not HasSelectionCriteria Then
        MsgBox "You can not run this function without Selection Criteria"
        Exit Function
End If
chkTerms = True

End Function


Private Sub chkSum_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkTermRpts_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
    If glbWFC Then
        If Index = 7 Then
            clpCode(7).TransDiv = GetTransDivTReason(clpCode(1).Text)
        End If
    End If
End Sub

Private Sub cmdPrintSelected_Click()
Dim x%

'On Error GoTo PrntErr

If Not chkTerms(1) Then Exit Sub

If chkTermRpts(0) = True Then GoTo Prt_OK
If chkTermRpts(1) = True Then GoTo Prt_OK
If chkTermRpts(2) = True Then GoTo Prt_OK
If chkTermRpts(3) = True Then GoTo Prt_OK
If chkTermRpts(4) = True Then GoTo Prt_OK
If chkTermRpts(5) = True Then GoTo Prt_OK

Exit Sub

Prt_OK:
x% = Cri_Select(1)        '0=View 1=Print
Screen.MousePointer = DEFAULT

Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"
Screen.MousePointer = DEFAULT

End Sub

Private Sub cmdTerminate_Click()
Dim TC As New ADODB.Recordset
Dim Msg$, DgDef As Variant, Response%
Dim xx1, xx2, xxx
Dim Emp_List As New Collection
Dim Num
Dim WSQLQ, SQLQ
Dim xStr
Dim recCount As Integer

xx2 = 0
If Not chkTerms(0) Then Exit Sub

Msg$ = Msg$ & "Are you sure you want to terminate "
Msg$ = Msg$ & Chr(10) & "these employees?"
Msg$ = Msg$ & Chr(10) & Chr(10) & "Make sure no other info:HR Window "
Msg$ = Msg$ & Chr(10) & "is open with these employees information showing."

Title$ = "Terminate Employees"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Modify
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " employee " Else Msg$ = Msg$ & " employees "
    Msg$ = Msg$ & "will be Terminated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No employee found for this seletion criteria to terminate."
    Exit Sub
End If

SQLQ = "SELECT ED_EMPNBR FROM HREMP "
SQLQ = SQLQ & " WHERE " & glbSeleDeptUn
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND ED_DIV='" & clpDiv.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_ORG='" & clpCode(0).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND ED_PT='" & clpPT.Text & "'"
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND ED_EMP='" & clpCode(2).Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC='" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(4).Text & "'"
If Len(clpCode(5).Text) > 0 Then SQLQ = SQLQ & " AND ED_ADMINBY='" & clpCode(5).Text & "'"
If Len(clpCode(6).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC='" & clpCode(6).Text & "'"
If Len(elpEEID.Text) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

TC.Open SQLQ, gdbAdoIhr001, adOpenKeyset
Do Until TC.EOF
    Emp_List.Add (TC("ED_EMPNBR"))
    TC.MoveNext
Loop

MailBody = ""
For Num = 1 To Emp_List.count
    glbLEE_ID = Emp_List(Num)
    lblEEID = Emp_List(Num)
    If gsEMAIL_ONTERM Then
        MailBody = MailBody & GetEmpName(glbLEE_ID) & vbCrLf
    End If
    If lblEEID > 0 Then
        Call SITerminate
        xx2 = xx2 + 1
    End If
Next

If gsEMAIL_ONTERM Then
     If Len(MailBody) > 0 Then
        If xx2 = 1 Then
            xStr = "The following employee has "
        Else
            xStr = "The following employees have "
        End If
        xStr = xStr & " been terminated." & vbCrLf
        xStr = xStr & "Termination Date: " & dlpTermDate & vbCrLf
        xStr = xStr & "Reason: " & GetTABLDesc("TERM", clpCode(1)) & vbCrLf & vbCrLf
        MailBody = xStr & MailBody
        Screen.MousePointer = DEFAULT
        Call imgEmail_Click
     End If
     
End If

If xx2 > 0 Then
    If xx2 = 1 Then
        MsgBox Val(xx2) & " employee is Terminated."
    Else
        MsgBox Val(xx2) & " employees are Terminated."
    End If
Else
    MsgBox "No Employees with this selection criteria exist!"
End If


glbLEE_ID = 0
dlpTermDate.Text = ""
clpCode(1).Text = ""
chkRehire.Value = True
txtComments.Text = ""
'cmdClose.SetFocus

Call UnloadFrms

MDIMain.panHelp(0).FloodType = 0

TC.Close

End Sub

Private Function GetEmpName(xempno)
Dim rsTemp As New ADODB.Recordset
Dim xStr, SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & xempno
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xStr = "Employee #:" & (xempno) & " Name: " & rsTemp("ED_FNAME") & " " & rsTemp("ED_SURNAME")
    End If
    rsTemp.Close
    GetEmpName = xStr
End Function

Public Sub imgEmail_Click()
Dim xEmail
On Error GoTo Email_Err
    If gsEMAIL_ONTERM Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetComPreferEmail("EMAIL_ONTERM")
        
        If Len(xEmail) > 0 Then
            frmSendEmail.txtTo.Text = xEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            'frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            frmSendEmail.txtSubject.Text = "info:HR Termination Notice"
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        Else
            MsgBox "There is no email for Email Notification on Salary on Company Preference screen. "
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

Private Sub cmdTerminate_GotFocus()
Call SetPanHelp(Me.ActiveControl) '19Aug99 js
End Sub

Private Sub CompTime()
Dim SQLQ, SQLX
On Error GoTo CompTm_Err


SQLQ = "DELETE FROM HRENTWRK " & in_SQL(glbIHRDBW) & " WHERE TE_WRKEMP='" & glbUserID & "'"

gdbAdoIhr001.Execute SQLQ

SQLQ = "SELECT ED_EMPNBR FROM HREMP "
SQLQ = SQLQ & " WHERE " & glbSeleDeptUn
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND ED_DIV='" & clpDiv.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_ORG='" & clpCode(0).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND ED_PT='" & clpPT.Text & "'"
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND ED_EMP='" & clpCode(2).Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC='" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(4).Text & "'"
If Len(clpCode(5).Text) > 0 Then SQLQ = SQLQ & " AND ED_ADMINBY='" & clpCode(5).Text & "'"
If Len(elpEEID.Text) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

SQLX = SQLQ

SQLQ = " SELECT '001' AS TE_COMPNO,AD_EMPNBR,'ADRE' AS TE_REASON_TABL,'CTOT' AS TE_REASON,"
If glbOracle Then
    SQLQ = SQLQ & " SUM(CASE WHEN SUBSTR(AD_REASON,1,2)='OT' THEN AD_HRS ELSE 0 END) AS TE_EARNHRS,"
    SQLQ = SQLQ & " SUM(CASE WHEN SUBSTR(AD_REASON,1,2)='CT' THEN AD_HRS ELSE 0 END) AS TE_USEDHRS"
ElseIf glbSQL Then
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS TE_EARNHRS,"
    SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS TE_USEDHRS"
Else
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='OT',AD_HRS,0)) AS TE_EARNHRS, "
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='CT',AD_HRS,0)) AS TE_USEDHRS "
End If
SQLQ = SQLQ & ",'" & glbUserID & "' AS TE_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE WHERE AD_EMPNBR IN (" & SQLX & ")"
SQLQ = SQLQ & " GROUP BY AD_COMPNO,AD_EMPNBR,AD_REASON_TABL "

SQLX = "INSERT INTO HRENTWRK "
SQLX = SQLX & "(TE_COMPNO,TE_EMPNBR,TE_REASON_TABL,TE_REASON,TE_EARNHRS,TE_USEDHRS,TE_WRKEMP)"
SQLX = SQLX & in_SQL(glbIHRDBW)
SQLX = SQLX & SQLQ
gdbAdoIhr001.Execute SQLX

Exit Sub

CompTm_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CompTime", "WORK File", "CREATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Sub


Private Sub Cri_EE()
Dim EECri As String

If Len(lblEEID) > 0 Then
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



Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_ORG"
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_LOC"
    Case 4: strCd$ = "HREMP.ED_REGION"
    Case 5: strCd$ = "HREMP.ED_ADMINBY"
    Case 6: strCd$ = "HREMP.ED_SHIFT"
    End Select
    CodeCri = "({" & strCd$ & "} = '" & clpCode(intIdx%).Text & "')"
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


Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True


End Sub
Private Function Cri_Select(RptTo As Integer)
Dim x%
Dim strWHand As String
Dim EECri As String
Dim glbstrSelCri1 As String

On Error GoTo CRW_Err

If RptTo = 1 Then
    If Not PrtForm("Termination Reports", Me) Then Exit Function
End If

Screen.MousePointer = HOURGLASS

'If Len(lblEEID) = 0 Then Exit Function
'glbstrSelCri = "{HREMP.ED_EMPNBR} = " & Val(lblEEID) & " "

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
For x% = 0 To 5
    If x% <> 1 Then
        Call Cri_Code(x%)
    End If
Next x%
Call Cri_PT
Call Cri_EE
glbiOneWhere = True
' reports names

If chkTermRpts(0).Value = True Then

    HisSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    HisSQL1 = " AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"

    Call SELATTWRK
    glbstrSelCri1 = glbstrSelCri & " AND {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzattdhs.rpt"
    Me.vbxCrystal.WindowTitle = "Attendance History Report"
    Me.vbxCrystal.Formulas(0) = "descGroup1 = 'TERMINATION'"
    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.SectionFormat(0) = "GH1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(2) = "GH2;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(3) = "GF2;F;X;X;X;X;X;X"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri1
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        'DBEngine.RegisterDatabase "IHR001", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & glbIHRDB & vbCr
        'Me.vbxCrystal.Connect = "DSN=IHR001;PWD=petman;DSQ=" & glbIHRDB
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5 + 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
        Me.vbxCrystal.DataFiles(5) = glbIHRDBW
    End If

    Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Termination :'"
    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
    Me.vbxCrystal.Formulas(2) = "DATERANGE = ''"
    Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If


If chkTermRpts(2).Value = True Then
    glbstrSelCri1 = glbstrSelCri & " AND {HREMPWRK.TT_WRKEMP}='" & glbUserID & "'"
    Call EmpWrk
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDBW
        For x% = 1 To 7
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
        'Me.vbxCrystal.Password = gstrAccPWord$
        'Me.vbxCrystal.UserName = gstrAccUID$
    End If
    Me.vbxCrystal.Formulas(51) = "showSIN = " & IIf(gSec_Show_SIN_SSN = 0, False, True) & " "
    Me.vbxCrystal.Formulas(52) = "showDOB = " & IIf(gSec_Show_DOB = 0, False, True) & " "
    Me.vbxCrystal.Formulas(53) = "showADDRESS = " & IIf(gSec_Show_ADDRESS = 0, False, True) & " "
    Me.vbxCrystal.Formulas(54) = "showMarital = " & IIf(gSec_Show_Marital = 0, False, True) & " "
        
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzprofil.rpt"
    Me.vbxCrystal.WindowTitle = "Employee Profile Report"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri1
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    'Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

If chkTermRpts(3).Value = True Then
    glbstrSelCri1 = glbstrSelCri & " AND {HRENTWRK.TE_WRKEMP}='" & glbUserID & "'"
    Call CompTime
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzenthr5.rpt"
    Me.vbxCrystal.WindowTitle = "Entitlements Report"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri1
    'Me.vbxCrystal.SectionFormat(0) = "GH1;F;X;X;X;X;X;X"
    'Me.vbxCrystal.SectionFormat(1) = "GF1;F;X;X;X;X;X;X"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 6
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
        Me.vbxCrystal.DataFiles(7) = glbIHRDBW
        ' set security for database
'        vbxCrystal.Password = gstrAccPWord$
'        vbxCrystal.UserName = gstrAccUID$
    End If
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

If chkTermRpts(4).Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzfollu1.rpt"
    Me.vbxCrystal.WindowTitle = lStr("Follow-ups Report")
    Me.vbxCrystal.Formulas(0) = "descGroup1 = 'TERMINATION'"
    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
        ' set security for database
'        vbxCrystal.Password = gstrAccPWord$
'        vbxCrystal.UserName = gstrAccUID$
    End If
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

If chkTermRpts(5).Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzmastr1.rpt"
    Me.vbxCrystal.WindowTitle = "Employee Comments"
    Me.vbxCrystal.Formulas(0) = "descGroup1 = 'TERMINATION'"
    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 5
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
        ' set security for database
'        vbxCrystal.Password = gstrAccPWord$
'        vbxCrystal.UserName = gstrAccUID$
    End If
    Me.vbxCrystal.Destination = RptTo
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    Me.vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

lblRptsPrinted.Visible = True

Exit Function

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

Resume Next

End Function
Private Sub EmpWrk()
Dim rsEmp As New ADODB.Recordset
Dim SQLX, SQLO
Dim SQLQ
Dim xEmplist
Dim xDate1, xDate2
On Error GoTo ERR_EmpWrk
xDate1 = DateAdd("yyyy", -100, Date)
xDate2 = DateAdd("yyyy", 50, Date)     'Jaddy 10/27/99

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0

SQLX = "DELETE FROM HREMPWRK " & in_SQL(glbIHRDBW) & " WHERE TT_WRKEMP='" & glbUserID & "'"

gdbAdoIhr001.Execute SQLX

SQLQ = "SELECT ED_EMPNBR FROM HREMP "
SQLQ = SQLQ & " WHERE " & glbSeleDeptUn
If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND ED_DIV='" & clpDiv.Text & "'"
If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_ORG='" & clpCode(0).Text & "'"
If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND ED_PT='" & clpPT.Text & "'"
If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND ED_EMP='" & clpCode(2).Text & "'"
If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC='" & clpCode(3).Text & "'"
If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(4).Text & "'"
If Len(clpCode(5).Text) > 0 Then SQLQ = SQLQ & " AND ED_ADMINBY='" & clpCode(5).Text & "'"
If Len(elpEEID.Text) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

If glbNoNONE Then
    SQLQ = SQLQ & " AND ED_ORG <> 'NONE' "
End If
If glbNoEXEC Then       'Hemu -EXE
    SQLQ = SQLQ & " AND ED_ORG <> 'EXEC' "
End If

rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsEmp.RecordCount < 1 Then
    MsgBox "NO RECORD SELECTED"
    GoTo rr
    Exit Sub
End If
xEmplist = ""
Do Until rsEmp.EOF
    xEmplist = xEmplist & "," & rsEmp("ED_EMPNBR")
    rsEmp.MoveNext
Loop
xEmplist = "(" & Mid(xEmplist, 2) & ")"
Call glbEmpWrk(xEmplist, xDate1, xDate2)

rr:
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""


Exit Sub
ERR_EmpWrk:
If Err = 13 Then
  MsgBox "SYSTEM ERROR : 13 - Type MisMatch"
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Create", "EMPWRK", "WORK FILE")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub dlpTermDate_LostFocus()
glbTermDate = dlpTermDate
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMUTERM"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMUTERM"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMUTERM"
glbchkSum = False  'Jaddy 11/9/99
Screen.MousePointer = HOURGLASS
Call setRptCaption(Me)

If glbLinamar Then
    clpCode(4).MaxLength = 8
End If
If glbCompSerial = "S/N - 2227W" Then
    clpCode(4).MaxLength = 6
End If

If glbWFC Then 'Ticket #15248
    lblTitle(0).Top = 330
    dlpTermDate.Top = 240
    lblTitle(1).Top = 690
    clpCode(1).Top = 630
    lblTitle(7).Top = 1020
    clpCode(7).Top = 990
    lblTitle(7).Visible = True
    clpCode(7).Visible = True
Else
    lblTermData.Top = 360
End If

MDIMain.panHelp(0).Caption = "Proceed with termination "

If Not gSec_Upd_Terminations Then
    chkRehire.Enabled = False
    chkSum.Enabled = False
    cmdTerminate.Enabled = False
    Panel3D1.Enabled = False
    panTermRpts.Enabled = False
    clpCode(0).Enabled = False
    clpCode(1).Enabled = False
    clpCode(2).Enabled = False
    clpCode(3).Enabled = False
    clpCode(4).Enabled = False
    clpCode(5).Enabled = False
    clpCode(6).Enabled = False
    
    txtComments.Enabled = False
    dlpTermDate.Enabled = False
End If
If Not gSec_Summarize_Attendance And glbLinamar Then chkSum.Enabled = False
Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmUTERM = Nothing 'carmen apr 2000
End Sub

Private Function InputHREMPEQU_DOT(EmpN As Long)
Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset

SQLQ = "SELECT * FROM HREMPEQU WHERE HREMPEQU.EQ_EMPNBR = "
SQLQ = SQLQ & EmpN


dynEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If dynEmp.RecordCount > 0 Then
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = " & Date_SQL(dlpTermDate.Text) & ", EQ_TYPE = 'T' "
    SQLQ = SQLQ & "WHERE HREMPEQU.EQ_EMPNBR = " & EmpN
    gdbAdoIhr001.Execute SQLQ
End If
dynEmp.Close
End Function

Private Function modTermAudit()
Dim x%, DtTm As Variant

Screen.MousePointer = HOURGLASS

modTermAudit = False

MDIMain.panHelp(0).FloodPercent = 50

MDIMain.panHelp(0).FloodPercent = 75
MDIMain.panHelp(0).FloodPercent = 100

modTermAudit = True
Screen.MousePointer = DEFAULT

Exit Function

Err_Msg:
Screen.MousePointer = DEFAULT
MsgBox "Problem Creating Audit record - Termination Aborted"

End Function

Private Function modTermMove()
Dim x%
Dim EEID&, TReason$, DtTm  As Variant, TRDesc$
Dim TComment$
Dim TRehire$
Dim TCause

Screen.MousePointer = HOURGLASS
modTermMove = False
DtTm = glbTermDate
EEID& = lblEEID
TReason$ = clpCode(1).Text
TComment$ = txtComments
TRehire$ = lblRehire
TRDesc$ = clpCode(1).Caption
TCause = clpCode(7).Text

x% = TERM_LIST(EEID&, DtTm, TReason$, TRDesc$, TComment$, TRehire$, TCause)
MDIMain.panHelp(0).FloodPercent = 5
x% = TERM_BASIC(EEID&)
MDIMain.panHelp(0).FloodPercent = 10
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EDUCSEM(EEID&)                  'laura nov 5, 1997
MDIMain.panHelp(0).FloodPercent = 13      '
If Not x Then GoTo modTermMoveErr_Msg    '
x% = TERM_ATTENDANCE(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 15
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_ATTENDANCE_HISTORY(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 20
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_JOB(EEID&)
MDIMain.panHelp(0).FloodPercent = 25
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_PERFORM(EEID&)
MDIMain.panHelp(0).FloodPercent = 30
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_SALARY(EEID&)
MDIMain.panHelp(0).FloodPercent = 35
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_HealthSafety(EEID&)
MDIMain.panHelp(0).FloodPercent = 38
x% = TERM_COMMENTS(EEID&)
MDIMain.panHelp(0).FloodPercent = 39
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_COBRA(EEID&)
MDIMain.panHelp(0).FloodPercent = 39
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_BENEFITS(EEID&)
MDIMain.panHelp(0).FloodPercent = 40
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_HealthCost(EEID&)
MDIMain.panHelp(0).FloodPercent = 40
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_OHS_Corrective(EEID&)
MDIMain.panHelp(0).FloodPercent = 40
If Not x Then GoTo modTermMoveErr_Msg
x% = Term_OHS_ROOT_CAUSES(EEID&)
MDIMain.panHelp(0).FloodPercent = 42
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_DOLENT(EEID&)

'Ticket #28789 - Actual Amounts Details
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_DOLENT_ACTDTL(EEID&)

MDIMain.panHelp(0).FloodPercent = 45
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_ENTHRS(EEID&)
MDIMain.panHelp(0).FloodPercent = 46
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EARN(EEID&)
MDIMain.panHelp(0).FloodPercent = 48
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EDU(EEID&)
MDIMain.panHelp(0).FloodPercent = 50
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EMPSKL(EEID&)
MDIMain.panHelp(0).FloodPercent = 52
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_TRADE(EEID&)
If Not x Then GoTo modTermMoveErr_Msg
If glbAxxent Then
x% = TERM_RSP(EEID&)                  'FRANK 12/22/2000
End If

x% = TERM_SUCCESSION(EEID&)          'George 04/04/2006 #10595
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_LANGUAGE(EEID&)          'George 04/04/2006 #10595
If Not x Then GoTo modTermMoveErr_Msg

x% = TERM_EMP_FLAGS(EEID&)          'Bryan 05/04/2006
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_GLDIST(EEID&)             'Bryan 05/04/2006
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EMPADP(EEID&)                  'FRANK 06/08/2006
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_EMPPAYROLL_TRANSACTION(EEID&)  'FRANK 03/18/2010 Ticket #18232
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_FOLLOW_UP(EEID&)  'Hemu 08/27/2010 Ticket #18668
If Not x Then GoTo modTermMoveErr_Msg
x% = TERM_HREEO(EEID&)  'Ticket #25669 Franks 06/24/2014
If Not x Then GoTo modTermMoveErr_Msg

If gsAttachment_DB Then
    x% = TERM_HRDOC_EMP(EEID&)                  'FRANK 01/10/2006
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_JOB_HISTORY(EEID&)          'George 01/19/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_COMMENTS(EEID&)          'George 01/26/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_HEALTH_SAFETY(EEID&)          'George 02/17/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_HEALTH_SAFETY_2(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_COUNSEL(EEID&)          'George 01/26/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_PERFORM_HISTORY(EEID&)          'George 01/26/2006 #10266
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_EDSEM(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_EDSEM_RETEST(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_HREDU(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
    x% = TERM_HRDOC_HRDOLENT(EEID&)
    If Not x Then GoTo modTermMoveErr_Msg
End If '

x% = InputHREMPEQU_DOT(EEID&)

modTermMove = True

Screen.MousePointer = DEFAULT
Exit Function

modTermMoveErr_Msg:
Screen.MousePointer = DEFAULT
MsgBox "Problem Creating Audit record - Termination Aborted"

End Function

Private Sub NukeEE2(EEID&)
Dim snapEETables As New ADODB.Recordset

Dim SQLQ As String, TabName$
Dim EEIDAlias$

On Error GoTo NukeEE2_Err
Dim rsSE As New ADODB.Recordset
Dim xUserID As String
rsSE.Open "SELECT USERID FROM HR_SECURE_BASIC WHERE EMPNBR=" & EEID&, gdbAdoIhr001, adOpenStatic
If Not rsSE.EOF Then
    xUserID = rsSE("USERID")
    Call NukeUSERID(xUserID)
End If
rsSE.Close

SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
'Ticket #22367, Ticket #20367 - Do not delete employee photo
SQLQ = SQLQ & " AND Table_Name <>'HR_PHOTO'"

'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
'SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"
'Ticket #20893 Franks 09/02/2011 - only remove data for the standard INFO:HR tables
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL IS NULL)"

snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEETables.RecordCount < 1 Then Exit Sub
snapEETables.MoveFirst

While Not snapEETables.EOF
    TabName$ = snapEETables("Table_Name")
    If UCase(Right(TabName$, 3)) <> "WRK" Then
      EEIDAlias$ = snapEETables("EMPNBR_Alias")
      Call NukeEERows2(TabName$, EEIDAlias$, EEID&)
    End If
    snapEETables.MoveNext
Wend
If glbAxxent Then '--
    TabName$ = "HRRSP"
    EEIDAlias$ = "RS_EMPNBR"
    Call NukeEERows2(TabName$, EEIDAlias$, EEID&)
End If
snapEETables.Close
Call UpdVacTimeRequest(EEID&, "D")
Exit Sub

NukeEE2_Err:
glbFrmCaption$ = "Delete Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_TABLES Error", "TabName$", "Search")
Call RollBack '29July99 js

End Sub

Private Sub NukeEERows2(TabName As String, EEIDAlias As String, EEID As Long)
' returns number of records found for ee in table
Dim Rows%, SQLQ As String

On Error GoTo NukeEERows2_Err

If TabName$ = "HREMPEQU" Then
    Exit Sub
End If

SQLQ = "DELETE FROM " & TabName
SQLQ = SQLQ & " WHERE " & EEIDAlias & " = " & EEID

gdbAdoIhr001.Execute SQLQ

Exit Sub

NukeEERows2_Err:
glbFrmCaption$ = "Nuke Rows"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete EE Rows", TabName$, "Delete")
Call RollBack '29July99 js

End Sub

Private Function ReadJob(IJob)
Dim rsTA1 As New ADODB.Recordset

ReadJob = "NO POSITION DESC - " & IJob


rsTA1.Open "HRJOB", gdbAdoIhr001, adOpenKeyset, adLockReadOnly, adCmdTableDirect

rsTA1.MoveFirst
rsTA1.Find "JB_CODE = '" & IJob & "'"

If rsTA1.EOF Then GoTo ReadJob_CLOSE 'Exit Function

ReadJob = rsTA1("JB_DESCR")

ReadJob_CLOSE:
    rsTA1.Close
End Function

Private Function READTABLE(Iname, Ikey)
Dim rsTA As New ADODB.Recordset

READTABLE = "No Table Description"


rsTA.Open "HRTABL", gdbAdoIhr001, adOpenKeyset, adLockReadOnly, adCmdTableDirect
rsTA.MoveFirst
rsTA.Filter = "TB_NAME = '" & Iname & "' and TB_KEY = '" & Ikey & "'"
If rsTA.EOF Then GoTo READTABLE_CLOSE 'Exit Function

READTABLE = rsTA("TB_DESC")
READTABLE_CLOSE:
rsTA.Close
End Function



Private Sub SELATTWRK()
Dim xlen, xxx, xx1
Dim db001 As Database
Dim SQLQ
Dim xFieldList

On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)
MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE "

If Len(HisSQL) > 1 Then
    SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
End If

MDIMain.panHelp(0).FloodPercent = 45
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)

MDIMain.panHelp(0).FloodPercent = 60
SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & " SELECT " & Replace(xFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
If Len(HisSQL1) > 1 Then
    SQLQ = SQLQ & "WHERE (" & HisSQL1 & ")"
End If
MDIMain.panHelp(0).FloodPercent = 75
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
If Not glbSQL And Not glbOracle Then Call Pause(1)

HisSQL = ""
HisSQL1 = ""
gdbAdoIhr001.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    gdbAdoIhr001.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    'If Err.Number = 2147217871 Then MsgBox Err.Description
    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
End Sub

Private Sub SITerminate()
Dim rsTB As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset

Dim Msg$, DgDef As Variant, Response%
Dim xx1, xxx
    MDIMain.panHelp(0).FloodType = 1
    rsTB.Open "Term_HRSEQ", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    
    If rsTB.EOF And rsTB.BOF Then
        glbTERM_Seq = 1
        rsTB.AddNew
    Else
        rsTB.MoveFirst 'Jaddy 10/28/99
        glbTERM_Seq = rsTB("TERM_SEQ_NEXT")
        'rsTB.Edit
    End If
    rsTB("TERM_SEQ_NEXT") = glbTERM_Seq + 1
    rsTB.Update
    
    rsTB.Close
    
    If Not AUDITTERM() Then MsgBox "ERROR - AUDIT FILE"
    Call UpdPaymentTypeVadim
    If Not modTermMove() Then Exit Sub
    If Not modTermAudit() Then Exit Sub
    
    EID& = CLng(lblEEID)
    TermDate$ = dlpTermDate.Text
    'Ticket #22367, Ticket #20367 - Do not delete employee photo
    'If glbSQL Or glbOracle Then
    '    gdbAdoIhr001.Execute "delete from HR_PHOTO where PT_EMPNBR=" & EID&
    'End If
    If Not Term_Superv() Then Exit Sub  'laura
    If Not Term_Reviewer() Then Exit Sub  'George Apr 4,2006 #10595
    
    MDIMain.panHelp(0).FloodPercent = 100
    
    Call NukeEE2(EID&)
    MDIMain.panHelp(0).FloodPercent = 0

    lblEEID = 0
    '~~~~~~~~~~~~~~~~~~~~~~~~~'added by RAUBREY 5/23/97 ~~~~~~~~~~~~~~~~~~~~~~
    
    rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If glbCompSerial = "S/N - 2436W" Then  'Family Day Ticket #27829
        rsT_PARCO("PC_NUMBER_EMPLOYEES") = modECount_FamilyDay
    Else
        rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsT_PARCO("PC_NUMBER_EMPLOYEES") - 1 'UPDATE FIELD WITH ACTUAL COUNT
    End If
    rsT_PARCO.Update
    rsT_PARCO.Close
End Sub

Private Sub UpdPaymentTypeVadim()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
If glbVadim Then
    If Vadim_PayType_field = "" Then Exit Sub
    SQLQ = "SELECT " & Vadim_PayType_field & " FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsEmp.EOF Then

    If glbChgTermReason = "LO" Or glbChgTermReason = "LAYO" Then
        rsEmp(Vadim_PayType_field) = "L"
        
        'City of Kawartha Lakes - When Terminate - they pass T code. If employee's benefits
        'are continued they do not terminate employee but instead just changes the Payment Type
        'R or L.
        If glbCompSerial = "S/N - 2363W" Then
            rsEmp(Vadim_PayType_field) = "T"
        End If
        
    ElseIf glbChgTermReason = "RETI" Then
        rsEmp(Vadim_PayType_field) = "R"
    
        'City of Kawartha Lakes - When Terminate - they pass T code. If employee's benefits
        'are continued they do not terminate employee but instead just changes the Payment Type
        'R or L.
        If glbCompSerial = "S/N - 2363W" Then
            rsEmp(Vadim_PayType_field) = "T"
        End If
    Else
        rsEmp(Vadim_PayType_field) = "T"
    End If
    rsEmp.Update
    End If
    rsEmp.Close
End If
End Sub

Private Function Term_Reviewer()
Dim SQLQDel As String, SQLQCom As String, strTable As String
Dim dynHRAT As New ADODB.Recordset
Dim strComm

On Error GoTo Database_Err
Term_Reviewer = False

strTable = "HR_SUCCESSION"
SQLQCom = "SELECT EU_REVIEWER FROM HR_SUCCESSION WHERE EU_REVIEWER = " & CLng(lblEEID)

dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

Screen.MousePointer = HOURGLASS

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        dynHRAT("EU_REVIEWER") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
End If
dynHRAT.Close

Screen.MousePointer = DEFAULT

Term_Reviewer = True
Exit Function

Database_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Reviewer", strTable, "TERMINATE")

End Function

Private Function Term_Superv()
'Laura
Term_Superv = False
Dim SQLQDel As String, SQLQCom As String, strTable As String
Dim dynHRAT As New ADODB.Recordset
Dim strComm

On Error GoTo Database_Err
'Set Superv_DB = OpenDatabase(glbIHRDB, False, False)

'select fields from HR_ATTENDANCE
strTable = "HR_ATTENDANCE"
SQLQCom = "SELECT * FROM HR_ATTENDANCE WHERE AD_SUPER = " & CLng(lblEEID)
dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

Screen.MousePointer = HOURGLASS

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        strComm = dynHRAT("AD_COMM")
        If strComm <> "" Then
            strComm = strComm & "; "
        End If
        'dynHRAT.Edit
        'dynHRAT("AD_COMM") = strComm & "Terminated Superviser: " & CLng(lblEEID) & "  " & RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
        dynHRAT("AD_SUPER") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
    dynHRAT.MoveFirst
End If
dynHRAT.Close

'select fields from HR_ATTENDANCE_HISTORY
strTable = "HR_ATTENDANCE_HISTORY"
SQLQCom = "SELECT * FROM HR_ATTENDANCE_HISTORY WHERE AH_SUPER = " & CLng(lblEEID)
dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        strComm = dynHRAT("AH_COMM")
        If strComm <> "" Then
            strComm = strComm & "; "
        End If
        'dynHRAT.Edit
        'dynHRAT("AH_COMM") = strComm & "Terminated Superviser: " & CLng(lblEEID) & "  " & RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
        dynHRAT("AH_SUPER") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
    dynHRAT.MoveFirst
End If

dynHRAT.Close
'select fields from HR_PERFORM_HISTORY
strTable = "HR_PERFORM_HISTORY"
SQLQCom = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_REPTAU = " & CLng(lblEEID)
dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        strComm = dynHRAT("PH_COMMENTS")
        If strComm <> "" Then
            strComm = strComm & "; "
        End If
        'dynHRAT.Edit
        'dynHRAT("PH_COMMENTS") = strComm & "Terminated Superviser: " & CLng(lblEEID) & "  " & RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
        dynHRAT("PH_REPTAU") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
    dynHRAT.MoveFirst
End If

dynHRAT.Close
'select fields from HR_JOB_HISTORY
strTable = "HR_JOB_HISTORY"
SQLQCom = "SELECT * FROM HR_JOB_HISTORY WHERE JH_REPTAU = " & CLng(lblEEID)
dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        'dynHRAT.Edit
        dynHRAT("JH_REPTAU") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
    dynHRAT.MoveFirst
End If
dynHRAT.Close
''select fields from HR_OCC_HEALTH_SAFETY
'strTable = "HR_OCC_HEALTH_SAFETY"
'SQLQCom = "SELECT * FROM HR_OCC_HEALTH_SAFETY WHERE EC_EMPNOT = " & CLng(lblEEID)
'Set dynHRAT = gdbIhr001.OpenRecordset(SQLQCom, dbOpenDynaset)
'If dynHRAT.RecordCount >= 1 Then
'    dynHRAT.MoveFirst
'    While Not dynHRAT.EOF
'        dynHRAT.Edit
'        dynHRAT("EC_EMPNOT") = 0
'        dynHRAT.Update
'        dynHRAT.MoveNext
'    Wend
'    dynHRAT.MoveFirst
'    dynHRAT.Close
'End If

'select fields from HR_OCC_HEALTH_SAFETY
strTable = "HR_OCC_HEALTH_SAFETY"
SQLQCom = "SELECT * FROM HR_OCC_HEALTH_SAFETY WHERE EC_EMPNOT = " & CLng(lblEEID)

dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        'dynHRAT.Edit
        dynHRAT("EC_EMPNOT") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
    dynHRAT.MoveFirst
End If
dynHRAT.Close
strTable = "HR_OCC_HEALTH_SAFETY"
SQLQCom = "SELECT * FROM HR_OCC_HEALTH_SAFETY WHERE EC_SUPERVISOR = " & CLng(lblEEID)

dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        'dynHRAT.Edit
        dynHRAT("EC_SUPERVISOR") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
    dynHRAT.MoveFirst
    dynHRAT.Close
End If
Screen.MousePointer = DEFAULT

Term_Superv = True
Exit Function

Database_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Superv", strTable, "TERMINATE")

End Function




Private Sub txtComments_GotFocus()

Call SetPanHelp(Me.ActiveControl)
MDIMain.panHelp(2).Caption = " "

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

Sub eeretrieve()
Call SET_UP_MODE
End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

'alpAPPNBR.Enabled = TF
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Terminations
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
Printable = False
End Property

Private Function getRecordCount_Modify()
    Dim SQLQ As String
    Dim rsEmp As New ADODB.Recordset
    Dim recCount As Integer
    Dim xEmplist As String
    
    getRecordCount_Modify = 0
    recCount = 0

    SQLQ = "SELECT COUNT(ED_EMPNBR) AS TOT_REC FROM HREMP "
    SQLQ = SQLQ & " WHERE " & glbSeleDeptUn
    If Len(clpDept.Text) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO='" & clpDept.Text & "'"
    If Len(clpDiv.Text) > 0 Then SQLQ = SQLQ & " AND ED_DIV='" & clpDiv.Text & "'"
    If Len(clpCode(0).Text) > 0 Then SQLQ = SQLQ & " AND ED_ORG='" & clpCode(0).Text & "'"
    If Len(clpPT.Text) > 0 Then SQLQ = SQLQ & " AND ED_PT='" & clpPT.Text & "'"
    If Len(clpCode(2).Text) > 0 Then SQLQ = SQLQ & " AND ED_EMP='" & clpCode(2).Text & "'"
    If Len(clpCode(3).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC='" & clpCode(3).Text & "'"
    If Len(clpCode(4).Text) > 0 Then SQLQ = SQLQ & " AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(4).Text & "'"
    If Len(clpCode(5).Text) > 0 Then SQLQ = SQLQ & " AND ED_ADMINBY='" & clpCode(5).Text & "'"
    If Len(clpCode(6).Text) > 0 Then SQLQ = SQLQ & " AND ED_LOC='" & clpCode(6).Text & "'"
    If Len(elpEEID.Text) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        recCount = rsEmp("TOT_REC")
    Else
        recCount = 0
    End If
    rsEmp.Close
    Set rsEmp = Nothing
    
    getRecordCount_Modify = recCount

End Function

