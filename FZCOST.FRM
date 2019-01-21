VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRCostE 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Cost of Employment"
   ClientHeight    =   10950
   ClientLeft      =   135
   ClientTop       =   1395
   ClientWidth     =   13110
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
   ScaleHeight     =   10950
   ScaleWidth      =   13110
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
      Left            =   11640
      TabIndex        =   36
      Tag             =   "Check to Include Indirect Salary"
      Top             =   3360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1995
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
      Left            =   2835
      TabIndex        =   14
      Tag             =   "00-Country of Employment"
      Top             =   4792
      Width           =   1440
   End
   Begin VB.Frame frmRPT 
      Height          =   2415
      Left            =   8520
      TabIndex        =   34
      Top             =   3120
      Width           =   3015
      Begin VB.CheckBox chkD1 
         Caption         =   "Total Entitlements"
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
         Left            =   240
         TabIndex        =   43
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkC2 
         Caption         =   "WSIB Absenteeism"
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
         Left            =   240
         TabIndex        =   42
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkC1 
         Caption         =   "O/S Entitlements"
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
         Left            =   240
         TabIndex        =   41
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkB2 
         Caption         =   "Company Paid Training/Seminars"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CheckBox chkB1 
         Caption         =   "Company Paid Associations"
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
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox chkA3 
         Caption         =   "Other Earnings"
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
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkA2 
         Caption         =   "Company Paid Benefits"
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
         Left            =   240
         TabIndex        =   37
         Top             =   480
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.CheckBox chkA1 
         Caption         =   "Base Salary"
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
         Left            =   240
         TabIndex        =   35
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2835
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "00-Employee Position Shift"
      Top             =   4458
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Tag             =   "EDPT-Category"
      Top             =   2788
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   2520
      TabIndex        =   12
      Tag             =   "00-Enter Section Code"
      Top             =   4124
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Tag             =   "00-Enter Position Group"
      Top             =   2454
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2520
      TabIndex        =   11
      Tag             =   "00-Enter Administered By Code"
      Top             =   3790
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2520
      TabIndex        =   10
      Tag             =   "00-Enter Region Code"
      Top             =   3456
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   2520
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1786
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1452
      Width           =   6915
      _ExtentX        =   12197
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
      Left            =   2520
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1118
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDLC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   784
      Width           =   6915
      _ExtentX        =   12197
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
      Left            =   2520
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   450
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   1
      MultiSelect     =   -1  'True
   End
   Begin Threed.SSOption optRtype 
      Height          =   225
      Index           =   0
      Left            =   180
      TabIndex        =   17
      Tag             =   "You can choose one of the options"
      Top             =   5820
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2831
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "  By Company"
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
   Begin Threed.SSOption optRtype 
      Height          =   225
      Index           =   1
      Left            =   1965
      TabIndex        =   18
      TabStop         =   0   'False
      Tag             =   "You can choose one of the option"
      Top             =   5835
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "  By Division"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optRtype 
      Height          =   225
      Index           =   2
      Left            =   3675
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "You can choose one of the options"
      Top             =   5850
      Width           =   1605
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "  By Department"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optRtype 
      Height          =   225
      Index           =   3
      Left            =   5475
      TabIndex        =   20
      TabStop         =   0   'False
      Tag             =   "You can choose one of the options"
      Top             =   5850
      Width           =   2055
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "  By Employee Name"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   4200
      TabIndex        =   9
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3122
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   8
      Tag             =   "40-Date from and including this date forward"
      Top             =   3122
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Tag             =   "10-Enter Employee Number"
      Top             =   2120
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6595
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6480
      Top             =   6360
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpAsOf 
      Height          =   285
      Left            =   2520
      TabIndex        =   15
      Tag             =   "40-Date from and including this date forward"
      Top             =   5160
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cost Salary Components As Of"
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
      Top             =   5205
      Width           =   2160
   End
   Begin VB.Label lblCountry 
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
      TabIndex        =   44
      Top             =   4863
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
      TabIndex        =   33
      Top             =   4527
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
      TabIndex        =   32
      Top             =   2847
      Width           =   630
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
      TabIndex        =   31
      Top             =   3855
      Width           =   1125
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
      TabIndex        =   30
      Top             =   4191
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
      TabIndex        =   29
      Top             =   2511
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
      TabIndex        =   28
      Top             =   3519
      Width           =   510
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
      Left            =   180
      TabIndex        =   27
      Top             =   1167
      Width           =   615
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
      TabIndex        =   26
      Top             =   3183
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
      TabIndex        =   25
      Top             =   2175
      Width           =   1290
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
      Left            =   180
      TabIndex        =   24
      Top             =   1839
      Width           =   450
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
      Left            =   180
      TabIndex        =   23
      Top             =   1503
      Width           =   840
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
      TabIndex        =   22
      Top             =   831
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
      TabIndex        =   21
      Top             =   495
      Width           =   555
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
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmRCostE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CodeCodes(7, 2) 'Lucy June 29, 2000
Dim NoDept As Long

Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc, xKeyPress, xTxtDiv, xLblDivDesc  'js-19Apr99
Dim LastlastID, LastlastNme, LastFirstNme, xTxtEEID, xLblEEName
Dim fglbUser

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Cost of Employment Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

PrntErr:
'MsgBox "Error Printing - check your Windows Printer setup"
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

Resume Next
Screen.MousePointer = DEFAULT

End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
'MsgBox "2"

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = True
glbstrSelCri = " {HRCOEWRK.WRKEMP}='" & glbUserID & "'"
If glbOracle Then
     Call Cri_SelectOracle
ElseIf glbSQL Then
    Call Cri_SelectSQL
Else
    Call Cri_Select
End If

'Ticket #30452 - ONE CARE - Direct / Show Indirect Salary
Me.vbxCrystal.Formulas(8) = "showIndirectSalary = " & IIf(chkIncIndSal, True, False) & " "

' report name
If optRtype(0).Value Then
    strRName$ = glbIHRREPORTS & "rzcost1.rpt"
ElseIf optRtype(1).Value Then
    strRName$ = glbIHRREPORTS & "rzcost2.rpt"
ElseIf optRtype(2).Value Then
    strRName$ = glbIHRREPORTS & "rzcost3.rpt"
ElseIf optRtype(3).Value Then
    strRName$ = glbIHRREPORTS & "rzcost4.rpt"
Else
    strRName$ = glbIHRREPORTS & "rzcost5.rpt"  'Lucy June 29, 2000
End If

Me.vbxCrystal.ReportFileName = strRName$
'If Not glbOracle Then
Me.vbxCrystal.Formulas(0) = "A1" & "= '" & IIf(chkA1, "Y", "N") & "'"
Me.vbxCrystal.Formulas(1) = "A2" & "= '" & IIf(chkA2, "Y", "N") & "'"
Me.vbxCrystal.Formulas(2) = "A3" & "= '" & IIf(chkA3, "Y", "N") & "'"
Me.vbxCrystal.Formulas(3) = "B1" & "= '" & IIf(chkB1, "Y", "N") & "'"
Me.vbxCrystal.Formulas(4) = "B2" & "= '" & IIf(chkB2, "Y", "N") & "'"
Me.vbxCrystal.Formulas(5) = "C1" & "= '" & IIf(chkC1, "Y", "N") & "'"
Me.vbxCrystal.Formulas(6) = "C2" & "= '" & IIf(chkC2, "Y", "N") & "'"
Me.vbxCrystal.Formulas(7) = "D1" & "= '" & IIf(chkD1, "Y", "N") & "'"
'End If

If glbNoNONE And glbNoEXEC Then     'Hemu -EXE
    If Len(glbstrSelCri) > 0 Then
        Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG }<> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
    Else
        Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR ({HREMP.ED_ORG } <> 'NONE' AND {HREMP.ED_ORG }<> 'EXEC'))"
    End If
ElseIf glbNoNONE Then
    If Len(glbstrSelCri) > 0 Then
        Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
    Else
        Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'NONE')"
    End If
ElseIf glbNoEXEC Then     'Hemu -EXE
    If Len(glbstrSelCri) > 0 Then
        Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'EXEC')"
    Else
        Me.vbxCrystal.SelectionFormula = "(isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG } <> 'EXEC')"
    End If
Else
    If glbCompSerial = "S/N - 2288W" Then 'Musashi - Ticket #12690
        'Limit the user has access to these employees salary information
        If Len(glbNoAccessGrp) <> 0 Then
            If Len(glbstrSelCri) > 0 Then
                glbstrSelCri = glbstrSelCri & " AND (" & glbNoAccessGrp & ") "
            Else
                glbstrSelCri = glbstrSelCri & " (" & glbNoAccessGrp & ") "
            End If
            glbstrSelCri = Replace(Replace(glbstrSelCri, "HREMP.ED_", "ED_"), "ED_", "HREMP.ED_")
            glbstrSelCri = Replace(glbstrSelCri, "(HREMP.", "({HREMP.")
            glbstrSelCri = Replace(glbstrSelCri, " HREMP.", " {HREMP.")
            glbstrSelCri = Replace(glbstrSelCri, " =", "} =")
        End If
    End If
    
    If Len(glbstrSelCri) > 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
End If

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
    For X% = 1 To 5
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next X%
    ' set security for database
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Cost of Employment"

Cri_SetAll = True


Screen.MousePointer = DEFAULT

Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cri_SetAll", "COE Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Function CriCheck()
Dim X%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 6
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If


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

If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text = "" Then
  MsgBox "If entering a date, both from and to date must be entered!"
  dlpDateRange(0).Text = ""
  dlpDateRange(0).SetFocus
  Exit Function
End If


If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then     'Hemu - 05/14/2003 - Check for blank
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If 'Hemu - 05/14/2003 - Check for blank
If dlpDateRange(1).Text <> "" And dlpDateRange(0).Text = "" Then
  MsgBox "If entering a date, both from and to date must be entered!"
  dlpDateRange(1).Text = ""
  dlpDateRange(1).SetFocus
  Exit Function
End If
If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True
End Function

Private Sub chkA1_Click()
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        If chkA1 Then
            chkIncIndSal.Visible = True
        Else
            chkIncIndSal.Visible = False
        End If
    Else
        chkIncIndSal.Visible = False
    End If

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    glbOnTop = "FRMRCOSTE"
    
    Screen.MousePointer = HOURGLASS
    
    If Not glbMulti Then
        lblShift.Visible = True
        txtShift.Visible = True
    End If
    
    Call setRptCaption(Me)
    
    optRtype(1).Caption = lStr(optRtype(1).Caption)
    optRtype(2).Caption = lStr(optRtype(2).Caption)
    optRtype(1).Caption = lStr("By Division")
    
    If glbLinamar Then clpCode(3).MaxLength = 8
    
    If glbSyndesis Then
        lblPosGroup.Caption = "Position Grade"
        clpCode(5).Tag = "00-Enter Position Grade"
    End If
    
    If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
    
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" Then
        chkIncIndSal.Visible = True
    Else
        chkIncIndSal.Value = False
        chkIncIndSal.Visible = False
    End If
    
    Call INI_Controls(Me)
    
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
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub


Private Sub optRtype_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Cri_SelectSQL()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ, SQLX
Dim xtime

On Error GoTo SelCri_Err

Screen.MousePointer = 11

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0
gdbAdoIhr001.Execute "DELETE FROM HRCOEWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"

MDIMain.panHelp(0).FloodPercent = 5

'---------HREMP
SQLQ = " SELECT '001' AS ED_COMPNO,ED_EMPNBR,ED_DIV, ED_DEPTNO,'EDOR' AS ED_ORG_TABL,ED_ORG,ED_PT,"
SQLQ = SQLQ & IIf(IsDate(dlpDateRange(0).Text), Date_SQL(dlpDateRange(0).Text), "Null") & " AS FROMDATE,"
SQLQ = SQLQ & IIf(IsDate(dlpDateRange(1).Text), Date_SQL(dlpDateRange(1).Text), "Null") & " AS TODATE"
SQLQ = SQLQ & ",'" & glbUserID & "' AS WRKEMP "
If clpCode(5).Text <> "" Then
    SQLQ = SQLQ & " FROM  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB=HRJOB.JB_CODE "
    SQLQ = SQLQ & " WHERE JB_GRPCD='" & clpCode(5).Text & "'"
    SQLQ = SQLQ & " AND JH_CURRENT<>0 AND "
Else
    SQLQ = SQLQ & " FROM HREMP WHERE "
End If
SQLQ = SQLQ & glbSeleDeptUn
If clpDiv.Text <> "" Then SQLQ = SQLQ & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then SQLQ = SQLQ & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then SQLQ = SQLQ & "AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then SQLQ = SQLQ & "AND ED_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')"
If clpCode(2).Text <> "" Then SQLQ = SQLQ & "AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then SQLQ = SQLQ & "AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then SQLQ = SQLQ & "AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(6).Text <> "" Then SQLQ = SQLQ & "AND ED_SECTION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(6).Text & "'"
If elpEEID.Text <> "" Then SQLQ = SQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then SQLQ = SQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"
If txtShift <> "" Then SQLQ = SQLQ & "AND ED_SHIFT='" & txtShift & "'"
If comCountryOfEmp.Text <> "" Then SQLQ = SQLQ & "AND ED_WORKCOUNTRY='" & comCountryOfEmp.Text & "'" 'Ticket #13621 Frank 09/07/2007
SQLQ = SQLQ & " ORDER BY ED_EMPNBR"

SQLX = "INSERT INTO HRCOEWRK (COMPANY,EMPLOYEE,DIVISION,DEPTNO,UNION_TABL,UNIONS,FT,FROMDATE,TODATE,WRKEMP) "
SQLX = SQLX & in_SQL(glbIHRDBW)
gdbAdoIhr001.Execute SQLX & SQLQ

MDIMain.panHelp(0).FloodPercent = 10

rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
xtime = (rsEmp.RecordCount / 500) * 3
If xtime < 2 Then xtime = 2
rsEmp.Close

'--------HR_SALARY_HISTORY
SQLQ = "SELECT SALS.SH_EMPNBR AS EMPNBR,SUM(SALS.SALANN) AS NUM1, SUM(SALS.SALRATE) AS NUM2, SUM(SALS.SH_WHRS) AS NUM3 "
SQLQ = SQLQ & " FROM " & SAL_SQL() & " AS SALS WHERE 1 = 1"
'Ticket #25196 - Adding As of Salary Date option
If Not IsDate(dlpAsOf.Text) Then
    SQLQ = SQLQ & " AND SALS.SH_CURRENT<>0 "
End If
If clpCode(5).Text <> "" Then SQLQ = SQLQ & " AND SALS.SH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpCode(5).Text & "')"
SQLQ = SQLQ & " GROUP BY SALS.SH_EMPNBR"

SQLX = " Update HRCOEWRK "
SQLX = SQLX & " SET SALARY = TRSNUM.NUM1, SALRATE=TRSNUM.NUM2, HRSPERWEEK=TRSNUM.NUM3"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

'TESTING
'-----------------------------------------------------------------------------------------------------------------------------------------
    'Ticket #30452 - ONE CARE - Direct / Indirect Salary Mass Update
    If glbCompSerial = "S/N - 2494W" And chkIncIndSal.Visible = True Then
        If chkIncIndSal Then
            MDIMain.panHelp(0).FloodPercent = 15
            
            'SQLQ = "SELECT SALS.SH_EMPNBR AS EMPNBR,SUM(SALS.SALANN) AS NUM1, SUM(SALS.SALRATE) AS NUM2, SUM(SALS.SH_WHRS) AS NUM3 "
            SQLQ = "SELECT SALS.SH_EMPNBR AS EMPNBR,SUM(SALS.SALANN) AS NUM1 "
            SQLQ = SQLQ & " FROM " & SAL1_SQL() & " AS SALS WHERE 1 = 1"
            'Ticket #25196 - Adding As of Salary Date option
            If Not IsDate(dlpAsOf.Text) Then
                SQLQ = SQLQ & " AND SALS.SH_CURRENT<>0 "
            End If
            If clpCode(5).Text <> "" Then SQLQ = SQLQ & " AND SALS.SH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpCode(5).Text & "')"
            SQLQ = SQLQ & " GROUP BY SALS.SH_EMPNBR"
            
            SQLX = " Update HRCOEWRK "
            SQLX = SQLX & " SET SALARY1 = TRSNUM.NUM1"
            SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
            gdbAdoIhr001.Execute SQLX
        End If
    End If
'-----------------------------------------------------------------------------------------------------------------------------------------

MDIMain.panHelp(0).FloodPercent = 20

'--------HRBENFT
SQLQ = "SELECT BF_EMPNBR AS EMPNBR, SUM(BF_CCOST) AS NUM1 FROM HRBENFT "
SQLQ = SQLQ & " WHERE BF_EDATE<=" & Date_SQL(IIf(IsDate(dlpDateRange(1).Text), dlpDateRange(1).Text, Date))
SQLQ = SQLQ & " GROUP BY BF_EMPNBR"

SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET BENEFITS = TRSNUM.NUM1"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 30

'--------HREARN
SQLQ = "SELECT EMPNBR,SUM(ACT_DOLLAR) AS NUM1 FROM HREARN WHERE COST_OF_EMPLOYMENT<>0"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " AND TDATE>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND FDATE<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET OTHEREARN = TRSNUM.NUM1"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX
   
MDIMain.panHelp(0).FloodPercent = 40

'--------HRTRADE
SQLQ = "SELECT TD_EMPNBR AS EMPNBR,SUM(TD_DUES) AS NUM1 FROM HRTRADE WHERE TD_COMPPD='Y'"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    'Changed by Frank 07/25/2001  TD_RENEWDT -> TD_BEGINDT
    SQLQ = SQLQ & " AND TD_BEGINDT>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND TD_BEGINDT<=" & Date_SQL(dlpDateRange(1).Text)
    'Changed by Frank 07/25/2001  TD_RENEWDT -> TD_BEGINDT
End If
SQLQ = SQLQ & " GROUP BY TD_EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET TRADES = TRSNUM.NUM1"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 50

'--------HREDSEM
SQLQ = "SELECT ES_EMPNBR AS EMPNBR,SUM(ES_TBCO) AS NUM1, SUM(ES_OTHER) AS NUM2, SUM(ES_ACCOM) AS NUM3 FROM HREDSEM "
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " WHERE ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY ES_EMPNBR"

SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET EDSEM = TRSNUM.NUM1+TRSNUM.NUM2+TRSNUM.NUM3"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 60

'---------HRDOLENT
SQLQ = "SELECT DE_EMPNBR AS EMPNBR,SUM(DE_ENTITLE) AS NUM1 ,SUM(DE_ACTUAL) AS NUM2 FROM HRDOLENT WHERE DE_COST_OF_EMPLOYMENT<>0"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " AND DE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND DE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY DE_EMPNBR"

SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET DEENTITLE=NUM1,DEDOLLARS = NUM1-NUM2"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 70

'-----------HRENTHRS
SQLQ = "SELECT HE_EMPNBR AS EMPNBR, SUM(HE_ENTITLE) AS NUM1, SUM(HE_TAKEN) AS NUM2 FROM HRENTHRS WHERE HE_COE<>0"
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " AND HE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND HE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY HE_EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET HRENTITLE = NUM1*SALRATE,OSHRENTITLE=(NUM1-NUM2)*SALRATE"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 80

'----------HR_ATTENDANCE
SQLQ = "SELECT AD_EMPNBR AS EMPNBR,"

SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS NUM1,"
SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS NUM2,"
SQLQ = SQLQ & " SUM(CASE WHEN LEFT(AD_REASON,3)='WCB' THEN AD_HRS ELSE 0 END) AS NUM3 "

SQLQ = SQLQ & " FROM HR_ATTENDANCE "
If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
    SQLQ = SQLQ & " WHERE AD_DOA>=" & Date_SQL(dlpDateRange(0).Text)
    SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpDateRange(1).Text)
End If
SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET OSCOMPTM = (NUM1-NUM2)*SALRATE,WCBDOL=NUM3*SALRATE"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN (" & SQLQ & ") AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

MDIMain.panHelp(0).FloodPercent = 90

'--------HREMP
SQLQ = "SELECT ED_EMPNBR AS EMPNBR,"
SQLQ = SQLQ & " ED_PVAC AS NUM1,ED_VAC AS NUM2,ED_PSICK AS NUM3,ED_SICK AS NUM4,ED_VACT AS NUM5,ED_SICKT AS NUM6"
SQLQ = SQLQ & " FROM HREMP"

SQLX = "Update HRCOEWRK "
SQLX = SQLX & " SET PVAC = NUM1*SALRATE,VAC=NUM2*SALRATE,PSICK=NUM3*SALRATE,SICK=NUM4*SALRATE,"
SQLX = SQLX & " OSVAC=(NUM1+NUM2-NUM5)*SALRATE,OSSICK = (NUM3+NUM4-NUM6)*SALRATE"
SQLX = SQLX & " FROM HRCOEWRK INNER JOIN  (" & SQLQ & ")  AS TRSNUM ON TRSNUM.EMPNBR=HRCOEWRK.EMPLOYEE "
gdbAdoIhr001.Execute SQLX

 
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET SALARY=0 WHERE SALARY IS NULL"

'Testing
'------------------------------------------------------------------------------------------------------------------------------------------------------
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET SALARY1=0 WHERE SALARY1 IS NULL"
'------------------------------------------------------------------------------------------------------------------------------------------------------

gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET HRSPERWEEK=0 WHERE SALRATE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET SALRATE=0 WHERE SALRATE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET BENEFITS=0 WHERE BENEFITS IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OTHEREARN=0 WHERE OTHEREARN IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET TRADES=0 WHERE TRADES IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET EDSEM=0 WHERE EDSEM IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSVAC=0 WHERE OSVAC IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSSICK=0 WHERE OSSICK IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET WCBDOL=0 WHERE WCBDOL IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSCOMPTM=0 WHERE OSCOMPTM IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET OSHRENTITLE=0 WHERE OSHRENTITLE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET DEDOLLARS=0 WHERE DEDOLLARS IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET DEENTITLE=0 WHERE DEENTITLE IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET PVAC=0 WHERE PVAC IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET VAC=0 WHERE VAC IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET PSICK=0 WHERE PSICK IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET SICK=0 WHERE SICK IS NULL"
gdbAdoIhr001W.Execute "UPDATE HRCOEWRK SET HRENTITLE=0 WHERE HRENTITLE IS NULL"
MDIMain.panHelp(0).FloodPercent = 100
   
'-------
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""

Exit Sub

SelCri_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cost of Employment", "HRCOEWRK", "INSERT")
Resume Next

End Sub

Private Function SAL_SQL()
Dim xStr
Dim strDays As String

'Commented by Bryan 19/Sep/05 - This is flakey code, replacing with something more solid
'If glbSQL Then
'    xStr = "(SELECT SH_EMPNBR,SH_ID,SH_SALARY,SH_SALCD,SH_WHRS,SH_CURRENT,SH_JOB, " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7)), HRPARCO.PC_DECHR) AS SALANN," & _
'    " (CASE WHEN HRPARCO.PC_SERIAL='S/N - 2236W' OR HRPARCO.PC_SERIAL='S/N - 2235W' THEN " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7))/ 52 / 5, HRPARCO.PC_DECHR) Else " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + ((CASE WHEN HR_SALARY_HISTORY.SH_WHRS=0 THEN 0 ELSE 1 / HR_SALARY_HISTORY.SH_WHRS/ 52 END) - 1) * (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7), HRPARCO.PC_DECHR) " & _
'    " END) AS SALRATE " & _
'    " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
'ElseIf glbOracle Then
'    xStr = "(SELECT SH_EMPNBR,SH_ID,SH_SALARY,SH_SALCD,SH_WHRS,SH_CURRENT,SH_JOB, " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7)), HRPARCO.PC_DECHR) AS SALANN," & _
'    " (CASE WHEN HRPARCO.PC_SERIAL='S/N - 2236W' OR HRPARCO.PC_SERIAL='S/N - 2235W' THEN " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7))/ 52 / 5, HRPARCO.PC_DECHR) Else " & _
'    " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + ((CASE WHEN HR_SALARY_HISTORY.SH_WHRS=0 THEN 0 ELSE 1 / HR_SALARY_HISTORY.SH_WHRS/ 52 END) - 1) * (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7), HRPARCO.PC_DECHR) " & _
'    " END) AS SALRATE " & _
'    " FROM HR_SALARY_HISTORY , HRPARCO WHERE HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
'Else
'    xStr = "QRY_SALARYS"
'End If

If GetLeapYear(Year(Date)) Then
    strDays = "366"
Else
    strDays = "355"
End If

xStr = "(SELECT SH_EMPNBR,SH_ID,SH_SALARY,SH_SALCD,SH_WHRS,SH_CURRENT,SH_JOB, "
If glbSQL Or glbOracle Then
    'Ticket #17654 - formula correction
    'If glbWFC Then
    '    xStr = xStr & "CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * 12), HRPARCO.PC_DECHR) " & _
    '     " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52 )), HRPARCO.PC_DECHR) " & _
    '     " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * " & strDays & "), HRPARCO.PC_DECHR) " & _
    '     " ELSE HR_SALARY_HISTORY.SH_SALARY END AS SALANN,"
    'Else
        xStr = xStr & "CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * 12), HRPARCO.PC_DECHR) " & _
         " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52 )), HRPARCO.PC_DECHR) " & _
         " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY / HR_JOB_HISTORY.JH_DHRS) * HR_SALARY_HISTORY.SH_WHRS * 52), HRPARCO.PC_DECHR) " & _
         " ELSE HR_SALARY_HISTORY.SH_SALARY END AS SALANN,"
    'End If
    'Calculation for Daily rate changed from this one to above - Ticket #17654
    '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * " & strDays & "), HRPARCO.PC_DECHR) " & _
    '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * 5 * 52), HRPARCO.PC_DECHR) " & _

    If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N = 2236W" Then
        xStr = xStr & " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7))/ 52 / 5, HRPARCO.PC_DECHR) Else " & _
        " ROUND(HR_SALARY_HISTORY.SH_SALARY * (1 + ((CASE WHEN HR_SALARY_HISTORY.SH_WHRS=0 THEN 0 ELSE 1 / HR_SALARY_HISTORY.SH_WHRS/ 52 END) - 1) * (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7), HRPARCO.PC_DECHR) " & _
        " END) AS SALRATE "
    Else
        'Ticket #17654 - formula correction
        'If glbWFC Then
        '    xStr = xStr & " CASE WHEN HR_SALARY_HISTORY.SH_WHRS =0 THEN 0 ELSE CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * 12) /HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
        '    " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND(HR_SALARY_HISTORY.SH_SALARY , HRPARCO.PC_DECHR) " & _
        '    " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52) , HRPARCO.PC_DECHR) " & _
        '    " ELSE ROUND(HR_SALARY_HISTORY.SH_SALARY* (1/HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
        '    " END END AS SALRATE "
        'Else
            xStr = xStr & " CASE WHEN HR_SALARY_HISTORY.SH_WHRS =0 THEN 0 ELSE CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * 12) /HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
            " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND(HR_SALARY_HISTORY.SH_SALARY , HRPARCO.PC_DECHR) " & _
            " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY / HR_JOB_HISTORY.JH_DHRS) , HRPARCO.PC_DECHR) " & _
            " ELSE ROUND(HR_SALARY_HISTORY.SH_SALARY* (1/HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
            " END END AS SALRATE "
        'End If
        
        'Calculation for Daily rate changed from this one to above - Ticket #17654
        '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52) , HRPARCO.PC_DECHR) " & _
        '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * 5) /HR_SALARY_HISTORY.SH_WHRS) , HRPARCO.PC_DECHR) " & _

    End If
Else
    xStr = xStr & "Format(iif(HR_SALARY_HISTORY.SH_SALCD='M',HR_SALARY_HISTORY.SH_SALARY * 12, " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52), " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='D',HR_SALARY_HISTORY.SH_SALARY * " & strDays & ", " & _
    "HR_SALARY_HISTORY.SH_SALARY))),Left('0.0000',(HRPARCO.PC_DECHR+2))) as SALANN, "
    If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N = 2236W" Then
        xStr = "QRY_SALARYS"
    Else
        xStr = xStr & "Format(iif(HR_SALARY_HISTORY.SH_WHRS=0,0,iif(HR_SALARY_HISTORY.SH_SALCD='M',((HR_SALARY_HISTORY.SH_SALARY * 12) /HR_SALARY_HISTORY.SH_WHRS/52), " & _
        "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY, " & _
        "iif(HR_SALARY_HISTORY.SH_SALCD='D',((HR_SALARY_HISTORY.SH_SALARY * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52), " & _
        "HR_SALARY_HISTORY.SH_SALARY* (1/HR_SALARY_HISTORY.SH_WHRS/52))))),Left('0.0000',(HRPARCO.PC_DECHR+2))) as SALRATE "
    End If
End If

If glbSQL Then
    'xStr = xStr & " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    
    xStr = xStr & " FROM ((HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    xStr = xStr & " INNER JOIN HR_JOB_HISTORY ON "
    'Ticket #25196 - Adding As of Salary Date option
    If Not IsDate(dlpAsOf.Text) Then
        xStr = xStr & " HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT AND"
    End If
    xStr = xStr & " HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
    
    If glbMulti Then
        xStr = xStr & " AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB"
    End If
    
    xStr = xStr & " )"
    
    'Ticket #25196 - Adding As of Salary Date option
    If glbMulti And Not IsDate(dlpAsOf.Text) Then
        xStr = xStr & " WHERE JH_POSITION_CONTROL = 'YES'"
    ElseIf IsDate(dlpAsOf.Text) Then
        'xStr = xStr & " WHERE SH_EDATE = (SELECT MAX(SH_EDATE) FROM HR_SALARY_HISTORY MSal WHERE MSal.SH_EMPNBR = HR_SALARY_HISTORY.SH_EMPNBR AND MSal.SH_EDATE <=" & Date_SQL(dlpAsOf.Text) & " GROUP BY MSal.SH_EMPNBR)"
        xStr = xStr & " WHERE SH_ID IN (SELECT SalID.MaxSalID FROM (SELECT MAX(SH_ID) AS MaxSalID, MaxSal.SH_EMPNBR FROM HR_SALARY_HISTORY MaxSal,"
        xStr = xStr & " (SELECT MAX(SH_EDATE) AS MaxEDate, MAX(SH_SDATE) AS MaxSDate, SH_EMPNBR FROM HR_SALARY_HISTORY WHERE SH_CURRENT = 0 AND SH_EDATE <=" & Date_SQL(dlpAsOf.Text) & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP)"
        xStr = xStr & " GROUP BY SH_EMPNBR) MaxSalEff"
        xStr = xStr & " WHERE MaxSal.SH_EMPNBR = MaxSalEff.SH_EMPNBR AND MaxSal.SH_EDATE = MaxSalEff.MaxEDate AND MaxSal.SH_SDATE = MaxSalEff.MaxSDate AND MaxSal.SH_CURRENT = 0"
        xStr = xStr & " GROUP BY MaxSal.SH_EMPNBR) SalID"
        xStr = xStr & " )"
    End If
    
    xStr = xStr & " ) "
    
ElseIf glbOracle Then
    'xStr = xStr & " FROM HR_SALARY_HISTORY,HRPARCO WHERE HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    
    xStr = xStr & " FROM HR_SALARY_HISTORY,HRPARCO,HR_JOB_HISTORY WHERE HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO "
    xStr = xStr & " AND HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"

ElseIf glbCompSerial <> "S/N - 2235W" And glbCompSerial <> "S/N = 2236W" Then
    'xStr = xStr & " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    
    xStr = xStr & " FROM ((HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    xStr = xStr & " INNER JOIN HR_JOB_HISTORY "
    xStr = xStr & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    xStr = xStr & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)) "
End If
'added by Bryan on 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    xStr = Replace(xStr, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
End If

SAL_SQL = xStr

End Function


Private Function SAL1_SQL()
Dim xStr
Dim strDays As String

If GetLeapYear(Year(Date)) Then
    strDays = "366"
Else
    strDays = "355"
End If

xStr = "(SELECT SH_EMPNBR,SH_ID,SH_SALARY1,SH_SALCD,SH_WHRS,SH_CURRENT,SH_JOB, "
If glbSQL Or glbOracle Then
    'Ticket #17654 - formula correction
    'If glbWFC Then
    '    xStr = xStr & "CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * 12), HRPARCO.PC_DECHR) " & _
    '     " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * (HR_SALARY_HISTORY.SH_WHRS * 52 )), HRPARCO.PC_DECHR) " & _
    '     " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * " & strDays & "), HRPARCO.PC_DECHR) " & _
    '     " ELSE HR_SALARY_HISTORY.SH_SALARY END AS SALANN,"
    'Else
        xStr = xStr & "CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY1 * 12), HRPARCO.PC_DECHR) " & _
         " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY1 * (HR_SALARY_HISTORY.SH_WHRS * 52 )), HRPARCO.PC_DECHR) " & _
         " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY1 / HR_JOB_HISTORY.JH_DHRS) * HR_SALARY_HISTORY.SH_WHRS * 52), HRPARCO.PC_DECHR) " & _
         " ELSE HR_SALARY_HISTORY.SH_SALARY1 END AS SALANN,"
    'End If
    'Calculation for Daily rate changed from this one to above - Ticket #17654
    '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * " & strDays & "), HRPARCO.PC_DECHR) " & _
    '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY * 5 * 52), HRPARCO.PC_DECHR) " & _

    If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N = 2236W" Then
        xStr = xStr & " ROUND(HR_SALARY_HISTORY.SH_SALARY1 * (1 + (HR_SALARY_HISTORY.SH_WHRS * 52 - 1)* (1 - (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7))/ 52 / 5, HRPARCO.PC_DECHR) Else " & _
        " ROUND(HR_SALARY_HISTORY.SH_SALARY1 * (1 + ((CASE WHEN HR_SALARY_HISTORY.SH_WHRS=0 THEN 0 ELSE 1 / HR_SALARY_HISTORY.SH_WHRS/ 52 END) - 1) * (72 - ASCII(HR_SALARY_HISTORY.SH_SALCD)) / 7), HRPARCO.PC_DECHR) " & _
        " END) AS SALRATE "
    Else
        'Ticket #17654 - formula correction
        'If glbWFC Then
        '    xStr = xStr & " CASE WHEN HR_SALARY_HISTORY.SH_WHRS =0 THEN 0 ELSE CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * 12) /HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
        '    " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND(HR_SALARY_HISTORY.SH_SALARY , HRPARCO.PC_DECHR) " & _
        '    " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52) , HRPARCO.PC_DECHR) " & _
        '    " ELSE ROUND(HR_SALARY_HISTORY.SH_SALARY* (1/HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
        '    " END END AS SALRATE "
        'Else
            xStr = xStr & " CASE WHEN HR_SALARY_HISTORY.SH_WHRS =0 THEN 0 ELSE CASE WHEN HR_SALARY_HISTORY.SH_SALCD='M' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY1 * 12) /HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
            " WHEN HR_SALARY_HISTORY.SH_SALCD='H' THEN ROUND(HR_SALARY_HISTORY.SH_SALARY1 , HRPARCO.PC_DECHR) " & _
            " WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND((HR_SALARY_HISTORY.SH_SALARY1 / HR_JOB_HISTORY.JH_DHRS) , HRPARCO.PC_DECHR) " & _
            " ELSE ROUND(HR_SALARY_HISTORY.SH_SALARY1* (1/HR_SALARY_HISTORY.SH_WHRS/52), HRPARCO.PC_DECHR) " & _
            " END END AS SALRATE "
        'End If
        
        'Calculation for Daily rate changed from this one to above - Ticket #17654
        '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52) , HRPARCO.PC_DECHR) " & _
        '" WHEN HR_SALARY_HISTORY.SH_SALCD='D' THEN ROUND(((HR_SALARY_HISTORY.SH_SALARY * 5) /HR_SALARY_HISTORY.SH_WHRS) , HRPARCO.PC_DECHR) " & _

    End If
Else
    xStr = xStr & "Format(iif(HR_SALARY_HISTORY.SH_SALCD='M',HR_SALARY_HISTORY.SH_SALARY1 * 12, " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY1 * (HR_SALARY_HISTORY.SH_WHRS * 52), " & _
    "iif(HR_SALARY_HISTORY.SH_SALCD='D',HR_SALARY_HISTORY.SH_SALARY1 * " & strDays & ", " & _
    "HR_SALARY_HISTORY.SH_SALARY1))),Left('0.0000',(HRPARCO.PC_DECHR+2))) as SALANN, "
    If glbCompSerial = "S/N - 2235W" Or glbCompSerial = "S/N = 2236W" Then
        xStr = "QRY_SALARYS"
    Else
        xStr = xStr & "Format(iif(HR_SALARY_HISTORY.SH_WHRS=0,0,iif(HR_SALARY_HISTORY.SH_SALCD='M',((HR_SALARY_HISTORY.SH_SALARY1 * 12) /HR_SALARY_HISTORY.SH_WHRS/52), " & _
        "iif(HR_SALARY_HISTORY.SH_SALCD='H',HR_SALARY_HISTORY.SH_SALARY1, " & _
        "iif(HR_SALARY_HISTORY.SH_SALCD='D',((HR_SALARY_HISTORY.SH_SALARY1 * " & strDays & ") /HR_SALARY_HISTORY.SH_WHRS/52), " & _
        "HR_SALARY_HISTORY.SH_SALARY1* (1/HR_SALARY_HISTORY.SH_WHRS/52))))),Left('0.0000',(HRPARCO.PC_DECHR+2))) as SALRATE "
    End If
End If

If glbSQL Then
    'xStr = xStr & " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    
    xStr = xStr & " FROM ((HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    xStr = xStr & " INNER JOIN HR_JOB_HISTORY ON "
    'Ticket #25196 - Adding As of Salary Date option
    If Not IsDate(dlpAsOf.Text) Then
        xStr = xStr & " HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT AND"
    End If
    xStr = xStr & " HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
    
    If glbMulti Then
        xStr = xStr & " AND HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB"
    End If
    
    xStr = xStr & " )"
    
    'Ticket #25196 - Adding As of Salary Date option
    If glbMulti And Not IsDate(dlpAsOf.Text) Then
        xStr = xStr & " WHERE JH_POSITION_CONTROL = 'YES'"
    ElseIf IsDate(dlpAsOf.Text) Then
        'xStr = xStr & " WHERE SH_EDATE = (SELECT MAX(SH_EDATE) FROM HR_SALARY_HISTORY MSal WHERE MSal.SH_EMPNBR = HR_SALARY_HISTORY.SH_EMPNBR AND MSal.SH_EDATE <=" & Date_SQL(dlpAsOf.Text) & " GROUP BY MSal.SH_EMPNBR)"
        xStr = xStr & " WHERE SH_ID IN (SELECT SalID.MaxSalID FROM (SELECT MAX(SH_ID) AS MaxSalID, MaxSal.SH_EMPNBR FROM HR_SALARY_HISTORY MaxSal,"
        xStr = xStr & " (SELECT MAX(SH_EDATE1) AS MaxEDate, MAX(SH_SDATE) AS MaxSDate, SH_EMPNBR FROM HR_SALARY_HISTORY WHERE SH_CURRENT = 0 AND SH_EDATE1 <=" & Date_SQL(dlpAsOf.Text) & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP)"
        xStr = xStr & " GROUP BY SH_EMPNBR) MaxSalEff"
        xStr = xStr & " WHERE MaxSal.SH_EMPNBR = MaxSalEff.SH_EMPNBR AND MaxSal.SH_EDATE1 = MaxSalEff.MaxEDate AND MaxSal.SH_SDATE = MaxSalEff.MaxSDate AND MaxSal.SH_CURRENT = 0"
        xStr = xStr & " GROUP BY MaxSal.SH_EMPNBR) SalID"
        xStr = xStr & " )"
    End If
    
    xStr = xStr & " ) "
    
ElseIf glbOracle Then
    'xStr = xStr & " FROM HR_SALARY_HISTORY,HRPARCO WHERE HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    
    xStr = xStr & " FROM HR_SALARY_HISTORY,HRPARCO,HR_JOB_HISTORY WHERE HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO "
    xStr = xStr & " AND HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)"

ElseIf glbCompSerial <> "S/N - 2235W" And glbCompSerial <> "S/N = 2236W" Then
    'xStr = xStr & " FROM HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    
    xStr = xStr & " FROM ((HR_SALARY_HISTORY INNER JOIN HRPARCO ON HR_SALARY_HISTORY.SH_COMPNO = HRPARCO.PC_CO) "
    xStr = xStr & " INNER JOIN HR_JOB_HISTORY "
    xStr = xStr & " ON HR_SALARY_HISTORY.SH_CURRENT = HR_JOB_HISTORY.JH_CURRENT"
    xStr = xStr & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR)) "
End If
'added by Bryan on 22/Sep/05 Ticket#9343
If glbCompSerial = "S/N - 2373W" Then 'Muskoka
    xStr = Replace(xStr, "SH_SALARY1", "SH_TOTAL", 1, , vbTextCompare)
End If

SAL1_SQL = xStr

End Function




Private Sub Cri_Select()
Dim rsEmp As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsORG As New ADODB.Recordset
Dim SQLQ, SQLX, xxx, xx1
Dim xEmpNbr, xtime, xSalRate, xNum1, xNum2, xNum3

On Error GoTo SelCri_Err

Screen.MousePointer = 11
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRCOEWRK  WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans
MDIMain.panHelp(0).FloodPercent = 5

gdbAdoIhr001W.BeginTrans
rsWRK.Open "HRCOEWRK", gdbAdoIhr001W, adOpenStatic, adLockOptimistic ' ,adCmdTableDirect

'---------HREMP
SQLQ = " SELECT '001' AS ED_COMPNO,ED_EMPNBR,ED_DIV, ED_DEPTNO,'EDOR' AS ED_ORG_TABL,ED_ORG,ED_PT,"
SQLQ = SQLQ & " ED_PVAC ,ED_VAC ,ED_PSICK,ED_SICK ,ED_VACT ,ED_SICKT, "
SQLQ = SQLQ & IIf(IsDate(dlpDateRange(0).Text), Date_SQL(dlpDateRange(0).Text), "Null") & " AS FROMDATE,"
SQLQ = SQLQ & IIf(IsDate(dlpDateRange(1).Text), Date_SQL(dlpDateRange(1).Text), "Null") & " AS TODATE"
If clpCode(5).Text <> "" Then
    SQLQ = SQLQ & " FROM  (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB=HRJOB.JB_CODE "
    SQLQ = SQLQ & " WHERE JB_GRPCD='" & clpCode(5).Text & "'"
    SQLQ = SQLQ & " AND JH_CURRENT<>0 AND "
Else
    SQLQ = SQLQ & " FROM HREMP WHERE "
End If
SQLQ = SQLQ & glbSeleDeptUn
If clpDiv.Text <> "" Then SQLQ = SQLQ & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then SQLQ = SQLQ & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then SQLQ = SQLQ & "AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then SQLQ = SQLQ & "AND ED_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')"
If clpCode(2).Text <> "" Then SQLQ = SQLQ & "AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then SQLQ = SQLQ & "AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then SQLQ = SQLQ & "AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(6).Text <> "" Then SQLQ = SQLQ & "AND ED_SECTION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(6).Text & "'"
If elpEEID.Text <> "" Then SQLQ = SQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then SQLQ = SQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"
If txtShift <> "" Then SQLQ = SQLQ & "AND ED_SHIFT='" & txtShift & "'"
If comCountryOfEmp.Text <> "" Then SQLQ = SQLQ & "AND ED_WORKCOUNTRY='" & comCountryOfEmp.Text & "'" 'Ticket #13621 Frank 09/07/2007
SQLQ = SQLQ & " ORDER BY ED_EMPNBR"
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
xxx = rsEmp.RecordCount
xx1 = 0
MDIMain.panHelp(0).FloodPercent = 10
Do Until rsEmp.EOF
    xEmpNbr = rsEmp("ED_EMPNBR")
    MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) * 0.8 + 10
    xx1 = xx1 + 1
    rsWRK.AddNew
    rsWRK("COMPANY") = "001"
    rsWRK("EMPLOYEE") = rsEmp("ED_EMPNBR")
    rsWRK("DIVISION") = rsEmp("ED_DIV")
    rsWRK("DEPTNO") = rsEmp("ED_DEPTNO")
    rsWRK("UNION_TABL") = "EDOR"
    rsWRK("UNIONS") = rsEmp("ED_ORG")
    rsWRK("FT") = rsEmp("ED_PT")
    rsWRK("FROMDATE") = rsEmp("FROMDATE")
    rsWRK("TODATE") = rsEmp("TODATE")
    '--------HR_SALARY_HISTORY
    SQLQ = "SELECT SUM(SALS.SALANN) AS NUM1, SUM(SALS.SALRATE) AS NUM2, SUM(SALS.SH_WHRS) AS NUM3 "
    SQLQ = SQLQ & " FROM " & SAL_SQL() & " AS SALS WHERE SALS.SH_CURRENT<>0 "
    If clpCode(5).Text <> "" Then SQLQ = SQLQ & " AND SALS.SH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpCode(5).Text & "')"
    SQLQ = SQLQ & " AND SALS.SH_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    If IsNull(rsORG("NUM3")) Then xNum3 = 0 Else xNum3 = rsORG("NUM3")
    rsWRK("SALARY") = xNum1
    rsWRK("SALRATE") = xNum2
    xSalRate = xNum2
    rsWRK("HRSPERWEEK") = xNum3
    rsORG.Close
    '--------HRBENFT
    SQLQ = "SELECT SUM(BF_CCOST) AS NUM1 FROM HRBENFT "
    SQLQ = SQLQ & " WHERE BF_EDATE<=" & Date_SQL(IIf(IsDate(dlpDateRange(1).Text), dlpDateRange(1).Text, Date))
    SQLQ = SQLQ & " AND BF_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    rsWRK("BENEFITS") = xNum1
    rsORG.Close
    '--------HREARN
    SQLQ = "SELECT SUM(ACT_DOLLAR) AS NUM1 FROM HREARN WHERE COST_OF_EMPLOYMENT<>0"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND TDATE>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND FDATE<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    SQLQ = SQLQ & " AND EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    rsWRK("OTHEREARN") = xNum1
    rsORG.Close
    '--------HRTRADE
    SQLQ = "SELECT SUM(TD_DUES) AS NUM1 FROM HRTRADE WHERE TD_COMPPD='Y'"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        'Changed by Frank 07/25/2001  TD_RENEWDT -> TD_BEGINDT
        SQLQ = SQLQ & " AND TD_BEGINDT>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND TD_BEGINDT<=" & Date_SQL(dlpDateRange(1).Text)
        'Changed by Frank 07/25/2001  TD_RENEWDT -> TD_BEGINDT
    End If
    SQLQ = SQLQ & " AND TD_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    rsWRK("TRADES") = xNum1
    rsORG.Close
    '--------HREDSEM
    SQLQ = "SELECT SUM(ES_TBCO) AS NUM1, SUM(ES_OTHER) AS NUM2, SUM(ES_ACCOM) AS NUM3 FROM HREDSEM "
    SQLQ = SQLQ & " WHERE ES_EMPNBR=" & xEmpNbr
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    If IsNull(rsORG("NUM3")) Then xNum3 = 0 Else xNum3 = rsORG("NUM3")
    rsWRK("EDSEM") = xNum1 + xNum2 + xNum3
    rsORG.Close
    '---------HRDOLENT
    SQLQ = "SELECT SUM(DE_ENTITLE) AS NUM1 ,SUM(DE_ACTUAL) AS NUM2 FROM HRDOLENT WHERE DE_COST_OF_EMPLOYMENT<>0"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND DE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND DE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    SQLQ = SQLQ & " AND DE_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    rsWRK("DEENTITLE") = xNum1
    rsWRK("DEDOLLARS") = xNum1 - xNum2
    rsORG.Close
    '-----------HRENTHRS
    
    SQLQ = "SELECT SUM(HE_ENTITLE) AS NUM1, SUM(HE_TAKEN) AS NUM2 FROM HRENTHRS WHERE HE_COE<>0"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND HE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND HE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    SQLQ = SQLQ & " AND HE_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    rsWRK("HRENTITLE") = xNum1 * xSalRate
    rsWRK("OSHRENTITLE") = (xNum1 - xNum2) * xSalRate
    rsORG.Close

    '----------HR_ATTENDANCE
    SQLQ = "SELECT "
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='OT', AD_HRS, 0)) AS NUM1,"
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,2)='CT', AD_HRS, 0)) AS NUM2,"
    SQLQ = SQLQ & " SUM(IIF(LEFT(AD_REASON,3)='WCB', AD_HRS, 0)) AS NUM3 "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR=" & xEmpNbr
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    If IsNull(rsORG("NUM3")) Then xNum3 = 0 Else xNum3 = rsORG("NUM3")
    rsWRK("OSCOMPTM") = (xNum1 - xNum2) * xSalRate
    rsWRK("WCBDOL") = xNum3 * xSalRate
    rsORG.Close
    '--------HREMP
    rsWRK("PVAC") = rsEmp("ED_PVAC") * xSalRate
    rsWRK("VAC") = rsEmp("ED_VAC") * xSalRate
    rsWRK("PSICK") = rsEmp("ED_PSICK") * xSalRate
    rsWRK("SICK") = rsEmp("ED_SICK") * xSalRate
    rsWRK("OSVAC") = (rsEmp("ED_PVAC") + rsEmp("ED_VAC") - rsEmp("ED_VACT")) * xSalRate
    rsWRK("OSSICK") = (rsEmp("ED_PSICK") + rsEmp("ED_SICK") - rsEmp("ED_SICKT")) * xSalRate
    rsWRK("WRKEMP") = glbUserID
    rsWRK.Update
    rsEmp.MoveNext
Loop
MDIMain.panHelp(0).FloodPercent = 100
rsEmp.Close
rsWRK.Close
gdbAdoIhr001W.CommitTrans
'-------
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
Exit Sub
SelCri_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cost of Employment", "HRCOEWRK", "INSERT")
Resume Next
End Sub

' Sam add for oracle
Private Sub Cri_SelectOracle()
Dim rsEmp As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsORG As New ADODB.Recordset
Dim SQLQ, SQLX, xxx, xx1
Dim xEmpNbr, xtime, xSalRate, xNum1, xNum2, xNum3

On Error GoTo SelCri_Err

Screen.MousePointer = 11

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 0

gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRCOEWRK  WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001W.CommitTrans
MDIMain.panHelp(0).FloodPercent = 5

gdbAdoIhr001W.BeginTrans
rsWRK.Open "HRCOEWRK", gdbAdoIhr001W, adOpenStatic, adLockOptimistic ' ,adCmdTableDirect

SQLQ = " SELECT '001' AS ED_COMPNO,ED_EMPNBR,ED_DIV, ED_DEPTNO,'EDOR' AS ED_ORG_TABL,ED_ORG,ED_PT,"
SQLQ = SQLQ & " ED_PVAC ,ED_VAC ,ED_PSICK,ED_SICK ,ED_VACT ,ED_SICKT, "
SQLQ = SQLQ & Date_SQL(dlpDateRange(0).Text) & " AS FROMDATE,"
SQLQ = SQLQ & Date_SQL(dlpDateRange(1).Text) & " AS TODATE"
If clpCode(5).Text <> "" Then
    SQLQ = SQLQ & " FROM  ( SELECT * FROM HREMP, HR_JOB_HISTORY,HRJOB WHERE HREMP.ED_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR AND HR_JOB_HISTORY.JH_JOB=HRJOB.JB_CODE )"
    SQLQ = SQLQ & " WHERE  JB_GRPCD='" & clpCode(5).Text & "'"
    SQLQ = SQLQ & " AND JH_CURRENT<>0 AND "
Else
    SQLQ = SQLQ & " FROM HREMP WHERE "
End If
SQLQ = SQLQ & glbSeleDeptUn
If clpDiv.Text <> "" Then SQLQ = SQLQ & " AND ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "')"
If clpDept.Text <> "" Then SQLQ = SQLQ & " AND ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "')"
If clpCode(0).Text <> "" Then SQLQ = SQLQ & "AND ED_LOC='" & clpCode(0).Text & "'"
If clpCode(1).Text <> "" Then SQLQ = SQLQ & "AND ED_ORG in ('" & Replace(clpCode(1).Text, ",", "','") & "')"
If clpCode(2).Text <> "" Then SQLQ = SQLQ & "AND ED_EMP in ('" & Replace(clpCode(2).Text, ",", "','") & "')"
If clpCode(3).Text <> "" Then SQLQ = SQLQ & "AND ED_REGION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(3).Text & "'"
If clpCode(4).Text <> "" Then SQLQ = SQLQ & "AND ED_ADMINBY='" & clpCode(4).Text & "'"
If clpCode(6).Text <> "" Then SQLQ = SQLQ & "AND ED_SECTION='" & IIf(glbLinamar, clpDiv.Text, "") & clpCode(6).Text & "'"
If elpEEID.Text <> "" Then SQLQ = SQLQ & "AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ")"
If clpPT.Text <> "" Then SQLQ = SQLQ & "AND ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"
If txtShift <> "" Then SQLQ = SQLQ & "AND ED_SHIFT='" & txtShift & "'"
If comCountryOfEmp.Text <> "" Then SQLQ = SQLQ & "AND ED_WORKCOUNTRY='" & comCountryOfEmp.Text & "'" 'Ticket #13621 Frank 09/07/2007
SQLQ = SQLQ & " ORDER BY ED_EMPNBR"
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
xxx = rsEmp.RecordCount
xx1 = 0
MDIMain.panHelp(0).FloodPercent = 10
Do Until rsEmp.EOF
    xEmpNbr = rsEmp("ED_EMPNBR")
    MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) * 0.8 + 10
    xx1 = xx1 + 1
    rsWRK.AddNew
    rsWRK("COMPANY") = "001"
    rsWRK("EMPLOYEE") = rsEmp("ED_EMPNBR")
    rsWRK("DIVISION") = rsEmp("ED_DIV")
    rsWRK("DEPTNO") = rsEmp("ED_DEPTNO")
    rsWRK("UNION_TABL") = "EDOR"
    rsWRK("UNIONS") = rsEmp("ED_ORG")
    rsWRK("FT") = rsEmp("ED_PT")
    rsWRK("FROMDATE") = rsEmp("FROMDATE")
    rsWRK("TODATE") = rsEmp("TODATE")
    '--------HR_SALARY_HISTORY
    SQLQ = "SELECT SUM(SALS.SALANN) AS NUM1, SUM(SALS.SALRATE) AS NUM2, SUM(SALS.SH_WHRS) AS NUM3 "
    SQLQ = SQLQ & " FROM " & SAL_SQL() & " SALS WHERE SALS.SH_CURRENT<>0 "
    If clpCode(5).Text <> "" Then SQLQ = SQLQ & " AND SALS.SH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpCode(5).Text & "')"
    SQLQ = SQLQ & " AND SALS.SH_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    If IsNull(rsORG("NUM3")) Then xNum3 = 0 Else xNum3 = rsORG("NUM3")
    rsWRK("SALARY") = xNum1
    rsWRK("SALRATE") = xNum2
    xSalRate = xNum2
    rsWRK("HRSPERWEEK") = xNum3
    rsORG.Close
    '--------HRBENFT
    SQLQ = "SELECT SUM(BF_CCOST) AS NUM1 FROM HRBENFT "
    If IsDate(dlpDateRange(1)) Then
        SQLQ = SQLQ & " WHERE BF_EDATE<=" & Date_SQL(dlpDateRange(1).Text)
    Else
        SQLQ = SQLQ & " WHERE BF_EDATE<=" & Date_SQL(Date)
    End If
    SQLQ = SQLQ & " AND BF_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    rsWRK("BENEFITS") = xNum1
    rsORG.Close
    '--------HREARN
    SQLQ = "SELECT SUM(ACT_DOLLAR) AS NUM1 FROM HREARN WHERE COST_OF_EMPLOYMENT<>0"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND TDATE>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND FDATE<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    SQLQ = SQLQ & " AND EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    rsWRK("OTHEREARN") = xNum1
    rsORG.Close
    '--------HRTRADE
    SQLQ = "SELECT SUM(TD_DUES) AS NUM1 FROM HRTRADE WHERE TD_COMPPD='Y'"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND TD_BEGINDT>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND TD_BEGINDT<=" & Date_SQL(dlpDateRange(1).Text)
        'Changed by Frank 07/25/2001  TD_RENEWDT -> TD_BEGINDT
    End If
    SQLQ = SQLQ & " AND TD_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    rsWRK("TRADES") = xNum1
    rsORG.Close
    '--------HREDSEM
    SQLQ = "SELECT SUM(ES_TBCO) AS NUM1, SUM(ES_OTHER) AS NUM2, SUM(ES_ACCOM) AS NUM3 FROM HREDSEM "
    SQLQ = SQLQ & " WHERE ES_EMPNBR=" & xEmpNbr
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    If IsNull(rsORG("NUM3")) Then xNum3 = 0 Else xNum3 = rsORG("NUM3")
    rsWRK("EDSEM") = xNum1 + xNum2 + xNum3
    rsORG.Close
    '---------HRDOLENT
    SQLQ = "SELECT SUM(DE_ENTITLE) AS NUM1 ,SUM(DE_ACTUAL) AS NUM2 FROM HRDOLENT WHERE DE_COST_OF_EMPLOYMENT<>0"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND DE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND DE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    SQLQ = SQLQ & " AND DE_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    rsWRK("DEENTITLE") = xNum1
    rsWRK("DEDOLLARS") = xNum1 - xNum2
    rsORG.Close
    '-----------HRENTHRS
    
    SQLQ = "SELECT SUM(HE_ENTITLE) AS NUM1, SUM(HE_TAKEN) AS NUM2 FROM HRENTHRS WHERE HE_COE<>0"
    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND HE_FDATE>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND HE_TDATE<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    SQLQ = SQLQ & " AND HE_EMPNBR=" & xEmpNbr
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    rsWRK("HRENTITLE") = xNum1 * xSalRate
    rsWRK("OSHRENTITLE") = (xNum1 - xNum2) * xSalRate
    rsORG.Close

    '----------HR_ATTENDANCE
    SQLQ = "SELECT "
    SQLQ = SQLQ & " SUM(AD_HRS) AS NUM1,"
    SQLQ = SQLQ & " SUM(AD_HRS) AS NUM2,"
    SQLQ = SQLQ & " SUM(AD_HRS) AS NUM3 "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " WHERE AD_EMPNBR=" & xEmpNbr
    SQLQ = SQLQ & "  AND SUBSTR(AD_REASON,1,2)='OT' AND SUBSTR(AD_REASON,1,2)='CT' AND SUBSTR(AD_REASON,1,3)='WCB' "

    If IsDate(dlpDateRange(1).Text) And IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND AD_DOA>=" & Date_SQL(dlpDateRange(0).Text)
        SQLQ = SQLQ & " AND AD_DOA<=" & Date_SQL(dlpDateRange(1).Text)
    End If
    rsORG.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If IsNull(rsORG("NUM1")) Then xNum1 = 0 Else xNum1 = rsORG("NUM1")
    If IsNull(rsORG("NUM2")) Then xNum2 = 0 Else xNum2 = rsORG("NUM2")
    If IsNull(rsORG("NUM3")) Then xNum3 = 0 Else xNum3 = rsORG("NUM3")
    rsWRK("OSCOMPTM") = (xNum1 - xNum2) * xSalRate
    rsWRK("WCBDOL") = xNum3 * xSalRate
    rsORG.Close
    '--------HREMP
    rsWRK("PVAC") = rsEmp("ED_PVAC") * xSalRate
    rsWRK("VAC") = rsEmp("ED_VAC") * xSalRate
    rsWRK("PSICK") = rsEmp("ED_PSICK") * xSalRate
    rsWRK("SICK") = rsEmp("ED_SICK") * xSalRate
    rsWRK("OSVAC") = (rsEmp("ED_PVAC") + rsEmp("ED_VAC") - rsEmp("ED_VACT")) * xSalRate
    rsWRK("OSSICK") = (rsEmp("ED_PSICK") + rsEmp("ED_SICK") - rsEmp("ED_SICKT")) * xSalRate
    rsWRK("WRKEMP") = glbUserID
    rsWRK.Update
    rsEmp.MoveNext
Loop
MDIMain.panHelp(0).FloodPercent = 100
rsEmp.Close
rsWRK.Close
gdbAdoIhr001W.CommitTrans
'-------
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
Exit Sub
SelCri_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cost of Employment", "HRCOEWRK", "INSERT")
Resume Next
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


