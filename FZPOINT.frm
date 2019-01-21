VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmRPoint 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Attendance Bonus Points Report"
   ClientHeight    =   8850
   ClientLeft      =   525
   ClientTop       =   1200
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
   ScaleHeight     =   8850
   ScaleWidth      =   12645
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkHideName 
      Caption         =   "Hide Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   35
      Top             =   6720
      Visible         =   0   'False
      Width           =   3075
   End
   Begin VB.TextBox txtIncentive 
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
      Left            =   3795
      MaxLength       =   1
      TabIndex        =   18
      Tag             =   "Incentive - Y or N"
      Top             =   4494
      Width           =   435
   End
   Begin VB.CheckBox chkPointType2 
      Caption         =   "L/LE Point"
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
      Left            =   6000
      TabIndex        =   29
      Top             =   5505
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox chkPointType1 
      Caption         =   "Absence Point"
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
      Left            =   4320
      TabIndex        =   28
      Top             =   5505
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CheckBox chkBF 
      Caption         =   "Bring Forward All Points"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   34
      Top             =   6345
      Width           =   3075
   End
   Begin VB.TextBox txtChargeCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2115
      MaxLength       =   15
      TabIndex        =   8
      Tag             =   "00-Enter Charge Code"
      Top             =   2904
      Width           =   855
   End
   Begin VB.CheckBox chkInclAtt 
      Caption         =   "Include Attendance History"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      TabIndex        =   33
      Top             =   5970
      Width           =   3075
   End
   Begin VB.TextBox txtPoint 
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
      Index           =   1
      Left            =   3090
      MaxLength       =   7
      TabIndex        =   22
      Tag             =   "10-Enter Hours"
      Top             =   5460
      Width           =   750
   End
   Begin VB.TextBox txtPoint 
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
      Index           =   0
      Left            =   2120
      MaxLength       =   7
      TabIndex        =   21
      Tag             =   "10-Enter Hours"
      Top             =   5460
      Width           =   750
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
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
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   678
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1314
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
      Left            =   1800
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   996
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1800
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   1950
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
      Index           =   2
      Left            =   1800
      TabIndex        =   10
      Tag             =   "ADRE-Attendance Reason"
      Top             =   3222
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "ADRE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3450
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3540
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   3540
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2268
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJOB 
      Height          =   285
      Left            =   7140
      TabIndex        =   23
      Tag             =   "00-Enter Position Code"
      Top             =   3855
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   10
      Left            =   1800
      TabIndex        =   20
      Tag             =   "00-Enter Administered By Code"
      Top             =   5130
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   9
      Left            =   1800
      TabIndex        =   19
      Tag             =   "00-Enter Region Code"
      Top             =   4812
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   8
      Left            =   7140
      TabIndex        =   27
      Tag             =   "00-Enter Section Code"
      Top             =   5130
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   7140
      TabIndex        =   24
      Tag             =   "00-Enter Position Group"
      Top             =   4170
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1800
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1632
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.TextBox txtShift 
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
      Left            =   2115
      MaxLength       =   4
      TabIndex        =   15
      Tag             =   "00-Shift individual is on"
      Top             =   4176
      Width           =   435
   End
   Begin VB.TextBox txtIncident 
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
      Left            =   2115
      MaxLength       =   1
      TabIndex        =   17
      Tag             =   "Incident - Y or N"
      Top             =   4494
      Width           =   435
   End
   Begin VB.TextBox txtHours 
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
      Left            =   7455
      MaxLength       =   7
      TabIndex        =   25
      Tag             =   "10-Enter Hours"
      Top             =   4494
      Width           =   750
   End
   Begin VB.TextBox txtSeniority 
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
      Left            =   7455
      MaxLength       =   1
      TabIndex        =   26
      Tag             =   "30-seniority Flag - Y or N"
      Top             =   4812
      Width           =   435
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
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   36
      Tag             =   "First Level of grouping records"
      Top             =   7080
      Width           =   2325
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
      Index           =   1
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Tag             =   "Second level of grouping records"
      Top             =   7425
      Width           =   2325
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
      Index           =   2
      Left            =   1770
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Tag             =   "Final sorting of records - no totals"
      Top             =   7770
      Width           =   2325
   End
   Begin VB.CheckBox chkVDesc 
      Caption         =   "View Attendance Descriptions"
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
      Left            =   6240
      TabIndex        =   39
      Top             =   3222
      Visible         =   0   'False
      Width           =   2895
   End
   Begin Threed.SSFrame fraCodeDesc 
      Height          =   1395
      Left            =   9480
      TabIndex        =   41
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   2461
      _StockProps     =   14
      Caption         =   "Attendance Code Descriptions"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
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
         Index           =   5
         Left            =   300
         TabIndex        =   45
         Top             =   975
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
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
         Index           =   4
         Left            =   285
         TabIndex        =   44
         Top             =   765
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
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
         Index           =   3
         Left            =   285
         TabIndex        =   43
         Top             =   555
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
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
         Index           =   2
         Left            =   285
         TabIndex        =   42
         Top             =   345
         Visible         =   0   'False
         Width           =   840
      End
   End
   Begin MSMask.MaskEdBox medDOW 
      Height          =   285
      Index           =   0
      Left            =   2120
      TabIndex        =   13
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   3858
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   0
      Left            =   1875
      TabIndex        =   30
      Tag             =   "Detailed Report"
      Top             =   6000
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Detailed"
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
   Begin MSMask.MaskEdBox medDOW 
      Height          =   285
      Index           =   1
      Left            =   2835
      TabIndex        =   14
      Tag             =   "10-Day of week 1=Sunday, 2=Monday..."
      Top             =   3858
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      MaxLength       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      Mask            =   "#"
      PromptChar      =   "_"
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   31
      TabStop         =   0   'False
      Tag             =   "Summary Report"
      Top             =   6000
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
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
   End
   Begin Threed.SSFrame fraDOW 
      Height          =   1905
      Left            =   9480
      TabIndex        =   46
      Top             =   240
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   3360
      _StockProps     =   14
      Caption         =   "Days of the Week"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "2 - Monday"
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
         Index           =   3
         Left            =   330
         TabIndex        =   53
         Top             =   525
         Width           =   960
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "3 - Tuesday"
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
         Index           =   4
         Left            =   330
         TabIndex        =   52
         Top             =   735
         Width           =   1020
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "4 - Wednesday"
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
         Index           =   5
         Left            =   330
         TabIndex        =   51
         Top             =   945
         Width           =   1290
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "5 - Thursday"
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
         Index           =   6
         Left            =   330
         TabIndex        =   50
         Top             =   1155
         Width           =   1080
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "6 - Friday"
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
         Index           =   7
         Left            =   330
         TabIndex        =   49
         Top             =   1365
         Width           =   810
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "7 - Saturday"
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
         Index           =   8
         Left            =   330
         TabIndex        =   48
         Top             =   1560
         Width           =   1050
      End
      Begin VB.Label lblDOW 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "1 - Sunday"
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
         Index           =   2
         Left            =   330
         TabIndex        =   47
         Top             =   320
         Width           =   930
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
      Top             =   7725
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
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   32
      TabStop         =   0   'False
      Tag             =   "Summary Report"
      Top             =   6000
      Width           =   1335
      _Version        =   65536
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "   Total Points"
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
   Begin INFOHR_Controls.CodeLookup clpChrgCode 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Specific Department Desired"
      Top             =   2904
      Visible         =   0   'False
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   16
      Tag             =   "00-Fund"
      Top             =   4176
      Visible         =   0   'False
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.EmployeeLookup elpSUP 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Tag             =   "00-Employee Number "
      Top             =   2586
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin VB.Label lblIncentive 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Incentive"
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
      Left            =   3000
      TabIndex        =   82
      Top             =   4539
      Width           =   660
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
      Left            =   90
      TabIndex        =   81
      Top             =   2625
      Width           =   945
   End
   Begin VB.Label lblPointFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Point"
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
      Left            =   90
      TabIndex        =   40
      Top             =   5505
      Width           =   1110
   End
   Begin VB.Label lblAttendCrit 
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
      Left            =   30
      TabIndex        =   80
      Top             =   120
      Width           =   1575
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
      Left            =   90
      TabIndex        =   79
      Top             =   405
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
      Left            =   90
      TabIndex        =   78
      Top             =   723
      Width           =   825
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
      Left            =   90
      TabIndex        =   77
      Top             =   1359
      Width           =   420
   End
   Begin VB.Label lblStatus 
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
      Left            =   90
      TabIndex        =   76
      Top             =   1677
      Width           =   450
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
      Left            =   90
      TabIndex        =   75
      Top             =   2313
      Width           =   1290
   End
   Begin VB.Label lblChargeCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Charge Code"
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
      Left            =   90
      TabIndex        =   74
      Top             =   2949
      Width           =   1410
   End
   Begin VB.Label lblAttCodes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance Codes"
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
      Left            =   90
      TabIndex        =   73
      Top             =   3267
      Width           =   1320
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
      Left            =   90
      TabIndex        =   72
      Top             =   3585
      Width           =   1095
   End
   Begin VB.Label lblDOW 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Day of Week"
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
      Left            =   90
      TabIndex        =   71
      Top             =   3903
      Width           =   945
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
      Left            =   90
      TabIndex        =   70
      Top             =   4221
      Width           =   1515
   End
   Begin VB.Label lblIncident 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Incident"
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
      Left            =   90
      TabIndex        =   69
      Top             =   4539
      Width           =   570
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
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
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   68
      Top             =   6000
      Width           =   1290
   End
   Begin VB.Label lblDOW 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "and"
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
      Left            =   2460
      TabIndex        =   67
      Top             =   3903
      Width           =   270
   End
   Begin VB.Label lblHours 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours ' >= '"
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
      Left            =   5940
      TabIndex        =   66
      Top             =   4539
      Width           =   795
   End
   Begin VB.Label lblSen1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Seniority"
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
      Left            =   5940
      TabIndex        =   65
      Top             =   4857
      Width           =   600
   End
   Begin VB.Label lblReprtGrping 
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
      TabIndex        =   64
      Top             =   6765
      Width           =   1695
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
      Left            =   150
      TabIndex        =   63
      Top             =   7125
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
      Left            =   150
      TabIndex        =   62
      Top             =   7455
      Width           =   885
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
      Left            =   150
      TabIndex        =   61
      Top             =   7770
      Width           =   660
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
      Left            =   90
      TabIndex        =   60
      Top             =   1041
      Width           =   615
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
      Left            =   90
      TabIndex        =   59
      Top             =   4857
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
      Left            =   90
      TabIndex        =   58
      Top             =   5175
      Width           =   1125
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
      Left            =   5940
      TabIndex        =   57
      Top             =   4221
      Visible         =   0   'False
      Width           =   1035
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
      Left            =   5940
      TabIndex        =   56
      Top             =   5175
      Width           =   540
   End
   Begin VB.Label lblJOB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   5940
      TabIndex        =   55
      Top             =   3903
      Visible         =   0   'False
      Width           =   975
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
      Left            =   90
      TabIndex        =   54
      Top             =   1995
      Width           =   630
   End
End
Attribute VB_Name = "frmRPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ReportSel, HisSQL, HisSQL1, HisSQLNoDate, HisSQL1NoDate
Dim WithEvents CN As ADODB.Connection
Attribute CN.VB_VarHelpID = -1

Private Sub chkInclAtt_Click()
    If Not gSec_Inq_Attendance_History And chkInclAtt.Value = 1 Then
        MsgBox "You Do Not Have Authority For This Transaction"
        chkInclAtt.Value = 0
    End If
End Sub

Private Sub chkPointType1_Click()
    If chkPointType1.Value <> 0 Then
        chkPointType2.Value = 0
    Else
        chkPointType2.Value = 1
    End If
End Sub

Private Sub chkPointType2_Click()
    If chkPointType2.Value <> 0 Then
        chkPointType1.Value = 0
    Else
        chkPointType1.Value = 1
    End If
End Sub

Private Sub chkVDesc_Click()
fraDOW.Visible = chkVDesc = 0
fraCodeDesc.Visible = Not chkVDesc = 0
If chkVDesc = 0 Then clpCode(2).SetFocus
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Attendance Bonus Points Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Public Sub cmdView_Click()
Dim x%, SQLQ
Dim strWHand As String, MyQuery As QueryDef
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    x% = Cri_SetAll()
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
'MsgBox Error$(Err)
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If MDIMain.panHelp(0).FloodType > 0 Then MDIMain.panHelp(0).FloodType = 0
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Private Sub clpChrgCode_Change()
txtChargeCode = clpChrgCode
End Sub

Private Sub clpCode_Change(Index As Integer)
    If Index = 3 Then
        txtShift = clpCode(3)
    End If
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

Call comGrpLoad

Call setRptCaption(Me)
Call setCaption(lblEENum(1))
Call setCaption(lblChargeCode)
Call setCaption(lblShift)
Call setCaption(lblHours)

If lblEENum(1).Caption = "AttSupervisor" Then lblEENum(1).Caption = "Supervisor"
lblAttCodes.Caption = lStr("Reason")
lblFromTo.Caption = lStr("From Date") & " / " & lStr("To Date")

If glbLinamar Then
    clpCode(9).MaxLength = 8
    lblDiv.FontBold = True
    lblFromTo.FontBold = True
End If

If glbMulti Then
    lblUnion.ForeColor = &HC000C0
    lblJOB.Visible = True
    clpJOB.Visible = True
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

'Casey House Ticket #5522
If glbCompSerial = "S/N - 2214W" Then
    lblChargeCode.Caption = "Attendance-Dept"
    lblShift.Caption = "Attendance-Fund"
    txtChargeCode.Visible = False
    clpChrgCode.Visible = True
    txtShift.MaxLength = 4
    txtShift.Visible = False
    clpCode(3).Visible = True
End If
If glbCompSerial = "S/N - 2396W" Then  'Oshawa CHC - Ticket #17323
    txtChargeCode.Visible = False
    lblChargeCode.Caption = lStr("G/L #")
    clpChrgCode.Visible = True
    clpChrgCode.LookupType = GL
    clpChrgCode.MaxLength = 25
    clpChrgCode.Tag = lStr("00-Enter G/L #")
End If
If glbBurlTech Then
    lblSen1.Caption = "Excused"
End If

If glbCompSerial = "S/N - 2214W" Then 'Casey House  - Ticket #15276
    lblIncentive.Caption = "HOOPP"
    txtIncentive.Tag = "HOOPP - Y or N"
End If

If Not gSec_Inq_Attendance_History Then
    chkInclAtt.Enabled = False
End If
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

Public Sub comGrpLoad()
    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("G/L")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem lStr("AttSupervisor")
    'Ticket #26118 - Add Account Code to G1 and G2
    comGroup(0).AddItem lStr("Account Code")
    
    comGroup(0).AddItem "(none)"
    
    
    comGroup(1).AddItem "Employee Name"
'    comGroup(1).AddItem "Attendance Code"
'    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem "Attendance Date"
    'Ticket #26118 - Add Account Code to G1 and G2
    comGroup(1).AddItem lStr("Account Code")
    
    
'    comGroup(2).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0


End Sub

Private Sub Cri_Attend()
Dim EECri As String, OneSet%, x%

OneSet% = False
If Len(clpCode(2).Text) = 0 Then Exit Sub

EECri = EECri & "'" & Replace(clpCode(2), ",", "','") & "' "

HisSQL = HisSQL & " AND HR_ATTENDANCE.AD_REASON IN (" & EECri & ")"
HisSQL1 = HisSQL1 & " AND HR_ATTENDANCE_HISTORY.AH_REASON IN (" & EECri & ")"

'Ticket #29107 - All selection criteria except From / To Date Range
HisSQLNoDate = HisSQLNoDate & " AND HR_ATTENDANCE.AD_REASON IN (" & EECri & ")"
HisSQL1NoDate = HisSQL1NoDate & " AND HR_ATTENDANCE_HISTORY.AH_REASON IN (" & EECri & ")"

EECri = "{HRATTWRK.AD_REASON} IN [" & EECri & "]"
If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

'OneSet% = False
'For X% = 2 To 5
'    If Len(clpCode(X%).Text) > 0 Then
'        OneSet% = OneSet% + 1
'    End If
'Next X%
'
'If OneSet% = 0 Then Exit Sub
'HisSQL = HisSQL & "AND ("
'HisSQL1 = HisSQL1 & "AND ("
'For X% = 2 To 5
'    If Len(clpCode(X%).Text) > 0 Then
'        HisSQL = HisSQL & "HR_ATTENDANCE.AD_REASON = '" & clpCode(X%).Text & "' "
'        HisSQL1 = HisSQL1 & "HR_ATTENDANCE_HISTORY.AH_REASON = '" & clpCode(X%).Text & "' "
'        EECri = EECri & "'" & clpCode(X%) & "' "
'        OneSet% = OneSet% - 1
'        If OneSet% > 0 Then
'            HisSQL = HisSQL & "OR "
'            HisSQL1 = HisSQL1 & "OR "
'            EECri = EECri & ","
'        Else
'            HisSQL = HisSQL & ") "
'            HisSQL1 = HisSQL1 & ") "
'        End If
'    End If
'Next X%
'EECri = "{HRATTWRK.AD_REASON} IN [" & EECri & "]"
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True


End Sub

Private Sub Cri_ChargeCode()
Dim EECri As String

If Len(txtChargeCode.Text) > 0 Then
    EECri = "{HRATTWRK.AD_CHRGCODE} = '" & txtChargeCode.Text & "' "
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

If Len(clpCode(intIdx%)) > 0 Then
    Select Case intIdx%
        Case 0: strCd$ = "HREMP.ED_LOC"
        Case 1: strCd$ = IIf(glbMulti, "HRATTWRK.AD_ORG", "HREMP.ED_ORG")
        Case 6: strCd$ = "HREMP.ED_EMP"
        Case 7: strCd$ = "HRJOB.JB_GRPCD"
        Case 8: strCd$ = "HREMP.ED_SECTION"  'Lucy June 30, 2000
        Case 9: strCd$ = "HREMP.ED_REGION"
        Case 10: strCd$ = "HREMP.ED_ADMINBY"
    End Select
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv & clpCode(intIdx%) & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%) & "') )"
    End If
End If

If glbMulti Then
    'Add by Franks Jan 29,2002 to fix the problem nothing show up on attendance report
    If intIdx% = 1 And Len(Trim(clpCode(1).Text)) > 0 Then
    'Add by Franks Jan 29,2002
        HisSQL = HisSQL & " AND AD_ORG='" & clpCode(intIdx%).Text & "' "
        HisSQL1 = HisSQL1 & " AND AH_ORG='" & clpCode(intIdx%).Text & "' "
        
        'Ticket #29107 - All selection criteria except From / To Date Range
        HisSQLNoDate = HisSQLNoDate & " AND AD_ORG='" & clpCode(intIdx%).Text & "' "
        HisSQL1NoDate = HisSQL1NoDate & " AND AH_ORG='" & clpCode(intIdx%).Text & "' "
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

Private Sub Cri_DOW()
Dim EECri As String, OneSet%
Dim x%

'If ReportSel = "ATT" And glbLinamar And optGrouping(0) Then
'    Exit Sub
'End If
For x% = 0 To 1
    If Len(medDOW(x%)) > 0 Then
        OneSet% = OneSet% + 1
    End If
Next x%

If OneSet% = 0 Then Exit Sub

For x% = 0 To 1
    If Len(medDOW(x%)) > 0 Then
        EECri = EECri & medDOW(x%)
        OneSet% = OneSet% - 1
    End If
    If OneSet% > 0 Then
        EECri = EECri & " ,"
    End If
Next x%
EECri = "{@DOANumDay} in [" & EECri & "]"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True


End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID) > 0 Then
    EECri = "{HREMP.ED_EMPNBR} IN [ " & getEmpnbr(elpEEID) & "] "
    HisSQL = HisSQL & " AND (HR_ATTENDANCE.AD_EMPNBR IN (" & getEmpnbr(elpEEID) & ")) "
    HisSQL1 = HisSQL1 & " AND (HR_ATTENDANCE_HISTORY.AH_EMPNBR IN (" & getEmpnbr(elpEEID) & ")) "
    
    'Ticket #29107 - All selection criteria except From / To Date Range
    HisSQLNoDate = HisSQLNoDate & " AND (HR_ATTENDANCE.AD_EMPNBR IN (" & getEmpnbr(elpEEID) & ")) "
    HisSQL1NoDate = HisSQL1NoDate & " AND (HR_ATTENDANCE_HISTORY.AH_EMPNBR IN (" & getEmpnbr(elpEEID) & ")) "
End If

If Len(EECri) >= 1 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND (" & EECri & ") "
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HRATTWRK.AD_DOA} "
    If glbSQL Then
        'Frank 07/13/04 Ticket# 6523 HRATTENDANCE -> HR_ATTENDANCE
        HisSQL = HisSQL & " AND ((HR_ATTENDANCE.AD_DOA) Between ('"
        HisSQL = HisSQL & Format(dlpDateRange(0), "mmm dd,yyyy") & "') And ('"
        HisSQL = HisSQL & Format(dlpDateRange(1), "mmm dd,yyyy") & "')) "
        HisSQL1 = HisSQL1 & " AND ((HR_ATTENDANCE_HISTORY.AH_DOA) Between ('"
        HisSQL1 = HisSQL1 & Format(dlpDateRange(0), "mmm dd,yyyy") & "') And ('"
        HisSQL1 = HisSQL1 & Format(dlpDateRange(1), "mmm dd,yyyy") & "')) "
    Else
        HisSQL = HisSQL & " AND ((HR_ATTENDANCE.AD_DOA) Between "
        HisSQL = HisSQL & Date_SQL(dlpDateRange(0).Text) & " And "
        HisSQL = HisSQL & Date_SQL(dlpDateRange(1).Text) & ") "
        HisSQL1 = HisSQL1 & " AND ((HR_ATTENDANCE_HISTORY.AH_DOA) Between "
        HisSQL1 = HisSQL1 & Date_SQL(dlpDateRange(0).Text) & " And "
        HisSQL1 = HisSQL1 & Date_SQL(dlpDateRange(1).Text) & ") "
    End If
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

For x% = 0 To 1
    If Len(dlpDateRange(x%).Text) > 0 Then
        TempCri = "({HRATTWRK.AD_DOA} "
        If x% = 0 Then
            TempCri = TempCri & " >= "
            HisSQL = HisSQL & " AND (HR_ATTENDANCE.AD_DOA >= " & Date_SQL(dlpDateRange(0)) & ") "
            HisSQL1 = HisSQL1 & " AND (HR_ATTENDANCE_HISTORY.AH_DOA >= " & Date_SQL(dlpDateRange(0)) & " ) "
        Else
            TempCri = TempCri & " <= "
            HisSQL = HisSQL & " AND (HR_ATTENDANCE.AD_DOA <= " & Date_SQL(dlpDateRange(1)) & " ) "
            HisSQL1 = HisSQL1 & " AND (HR_ATTENDANCE_HISTORY.AH_DOA <= " & Date_SQL(dlpDateRange(1)) & " ) "
        End If
        dtYYY% = Year(dlpDateRange(x%).Text)
        dtMM% = month(dlpDateRange(x%).Text)
        dtDD% = Day(dlpDateRange(x%).Text)
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

Private Sub Cri_Points0()
Dim EECri As String

If Len(txtPoint(0).Text) > 0 Then
     EECri = "{HRATTWRK.AD_WHRS} >= " & Val(txtPoint(0).Text) & " "
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

Private Sub Cri_Points1()
Dim EECri As String

If Len(txtPoint(1).Text) > 0 Then
     EECri = "{HRATTWRK.AD_WHRS} <= " & Val(txtPoint(1).Text) & " "
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

Private Sub Cri_Hours()
Dim EECri As String

If Len(txtHours.Text) > 0 Then
    EECri = "{HRATTWRK.AD_HRS} >= " & Val(txtHours.Text) & " "
    HisSQL = HisSQL & " AND (HR_ATTENDANCE.AD_HRS = " & Val(txtHours.Text) & ") "
    HisSQL1 = HisSQL1 & " AND (HR_ATTENDANCE_HISTORY.AH_HRS = " & Val(txtHours.Text) & ") "
        
    'Ticket #29107 - All selection criteria except From / To Date Range
    HisSQLNoDate = HisSQLNoDate & " AND (HR_ATTENDANCE.AD_HRS = " & Val(txtHours.Text) & ") "
    HisSQL1NoDate = HisSQL1NoDate & " AND (HR_ATTENDANCE_HISTORY.AH_HRS = " & Val(txtHours.Text) & ") "
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

Private Sub Cri_Incident()
Dim EECri As String

If Len(txtIncident.Text) > 0 Then
        If txtIncident = "Y" Then
            EECri = "{HRATTWRK.AD_INCID}"
        Else
            EECri = " NOT {HRATTWRK.AD_INCID}"
        End If
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

Private Sub Cri_Incentive()
Dim EECri As String

If Len(txtIncentive.Text) > 0 Then
    If txtIncentive = "Y" Then
        EECri = "{HRATTWRK.AD_INDICATOR}"
    Else
        EECri = " NOT {HRATTWRK.AD_INDICATOR}"
    End If
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

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_Seniority()
Dim EECri As String

If Len(txtSeniority.Text) > 0 Then
    If txtSeniority = "Y" Then
        EECri = "{HRATTWRK.AD_SEN}"
    Else
        EECri = " NOT {HRATTWRK.AD_SEN}"
    End If
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
Dim x%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_Code(0)
Call Cri_Code(6)
Call Cri_Code(8)
Call Cri_Code(9)
Call Cri_Code(10)
Call Cri_PT

If glbMulti Then
    HisSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    HisSQL1 = " AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    
    'Ticket #29107 - All selection criteria except From / To Date Range
    HisSQLNoDate = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    HisSQL1NoDate = " AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    
    Call Cri_Code(1)
Else
    Call Cri_Code(1)
    HisSQL = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    HisSQL1 = " AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    
    'Ticket #29107 - All selection criteria except From / To Date Range
    HisSQLNoDate = " AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
    HisSQL1NoDate = " AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"
End If

Call Cri_EE
Call Cri_ChargeCode   'laura 03/03/98
Call Cri_Hours
Call Cri_Attend
Call Cri_FTDates
Call Cri_DOW
Call Cri_Shift
Call Cri_Sup    'supervisor
Call Cri_Incident
Call Cri_Incentive  'Ticket #15276
Call Cri_Seniority
Call Cri_Job
Call Cri_Code(7)
'Call Cri_Points0
'Call Cri_Points1

glbstrSelCri = glbstrSelCri & " AND {HRATTWRK.AD_WRKEMP}='" & glbUserID & "'"

Call SELATTWRK

If glbBurlTech Then
    If optGrouping(0) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2351pd.rpt"
    ElseIf optGrouping(1) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2351ps.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "sn2351pt.rpt"
    End If
Else
    If optGrouping(0) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzpoint.rpt"
    ElseIf optGrouping(1) Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzpoints.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzpointt.rpt"
    End If
    
    'Release 8.0 - Ticket #22682: Hide Name
    If chkHideName.Visible Then
        Me.vbxCrystal.Formulas(10) = "hideName=" & (chkHideName <> 0)
    End If
End If
Me.vbxCrystal.WindowTitle = "Attendance Report"

x% = Cri_Sorts()   ' returns number of sections formated

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 10
        If x% <> 3 Then
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Else
            Me.vbxCrystal.DataFiles(3) = glbIHRDBW
        End If
    Next x%
    If optGrouping(0) Then
        Me.vbxCrystal.DataFiles(11) = glbIHRDB
    End If
End If

Cri_SetAll = True


Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Cri_SetAll = False
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Cri_Shift()
Dim EECri As String

If Len(txtShift.Text) > 0 Then
    EECri = "{HRATTWRK.AD_SHIFT} = '" & txtShift.Text & "' "
    HisSQL = HisSQL & " AND (HR_ATTENDANCE.AD_SHIFT = '" & txtShift.Text & "') "
    HisSQL1 = HisSQL1 & " AND (HR_ATTENDANCE_HISTORY.AH_SHIFT = '" & txtShift.Text & "') "
    
    'Ticket #29107 - All selection criteria except From / To Date Range
    HisSQLNoDate = HisSQLNoDate & " AND (HR_ATTENDANCE.AD_SHIFT = '" & txtShift.Text & "') "
    HisSQL1NoDate = HisSQL1NoDate & " AND (HR_ATTENDANCE_HISTORY.AH_SHIFT = '" & txtShift.Text & "') "
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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$, DOA
Dim dscGroup$, GrpIdx%, SavGrp1, SavGrp2, SavGrp3
Dim strSMonth$, strSPoint$

z% = 0
Cri_Sorts = 0
If dlpDateRange(0) <> "" And dlpDateRange(1) <> "" Then
    strSFormat$ = "As of " & dlpDateRange(0) & " through " & dlpDateRange(1)
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
    strSMonth$ = DateDiff("m", dlpDateRange(0), dlpDateRange(1)) + 1
    Me.vbxCrystal.Formulas(3) = "Mths = " & strSMonth$
ElseIf dlpDateRange(0) <> "" Then
    strSFormat$ = "As From " & dlpDateRange(0)
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
    strSMonth$ = DateDiff("m", dlpDateRange(0), CVDate(Format("12/31/9999", "mm/dd/yyyy"))) + 1
    Me.vbxCrystal.Formulas(3) = "Mths = " & strSMonth$
ElseIf dlpDateRange(1) <> "" Then
    strSFormat$ = "Up To " & dlpDateRange(1)
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
    strSMonth$ = DateDiff("m", CVDate(Format("1/1/1900", "mm/dd/yyyy")), dlpDateRange(1)) + 1
    Me.vbxCrystal.Formulas(3) = "Mths = " & strSMonth$
Else
    strSFormat$ = "No date entered"
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"
    strSMonth$ = "1"
    Me.vbxCrystal.Formulas(3) = "Mths = " & strSMonth$
End If

If txtPoint(0) <> "" And txtPoint(1) <> "" Then
    strSPoint$ = "Points from " & txtPoint(0) & " to " & txtPoint(1)
    Me.vbxCrystal.Formulas(4) = "PointRange = '" & strSPoint$ & "'"
Else
    strSPoint$ = ""
    Me.vbxCrystal.Formulas(4) = "PointRange = '" & strSPoint$ & "'"
End If

If comGroup(0).Text = lStr("AttSupervisor") Then
    grpField$ = "{@fldADSuper}"
Else
    grpField$ = getEGroup(comGroup(0).Text)
End If

If grpField$ = "(none)" Then grpField$ = "{HREMP.ED_COMPNO}"
'Added by Bryan 30/May/05 Ticket#10902

SavGrp1 = grpField$

If Not (grpField$ = "{HREMP.ED_COMPNO}") Then 'If GrpIdx% < 5 Then
  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
Else
  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
  Me.vbxCrystal.SectionFormat(z%) = "GF1;F;X;X;X;X;X;X"
  z% = z% + 1
End If
Me.vbxCrystal.GroupCondition(0) = "GROUP1;" & grpField$ & ";ANYCHANGE;A"

Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpSUP(1)) > 0 Then
    'EECri = "{HRATTWRK.AD_SUPER} IN (" & getEmpnbr(elpSUP(1)) & ") "
    HisSQL = HisSQL & " AND (HR_ATTENDANCE.AD_SUPER IN (" & getEmpnbr(elpSUP(1)) & ")) "
    HisSQL1 = HisSQL1 & " AND (HR_ATTENDANCE_HISTORY.AH_SUPER IN (" & getEmpnbr(elpSUP(1)) & ")) "
    
    'Ticket #29107 - All selection criteria except From / To Date Range
    HisSQLNoDate = HisSQLNoDate & " AND (HR_ATTENDANCE.AD_SUPER IN (" & getEmpnbr(elpSUP(1)) & ")) "
    HisSQL1NoDate = HisSQL1NoDate & " AND (HR_ATTENDANCE_HISTORY.AH_SUPER IN (" & getEmpnbr(elpSUP(1)) & ")) "
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

Private Function CriCheck()
Dim x%

CriCheck = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv) > 0 And clpDiv.Caption = "Unassigned" Then
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

If elpSUP(1).Caption = "Unnasigned" Then
    MsgBox "If Employee Entered - they must exist"
    elpSUP(1).SetFocus
    Exit Function
End If

For x% = 0 To 10
    If x% <> 4 And x% <> 5 Then
     If Not clpCode(x).ListChecker Then Exit Function
    End If
Next x%

For x% = 0 To 1
 If Len(dlpDateRange(x%)) > 0 Then
    If Not IsDate(dlpDateRange(x%)) Then
        MsgBox "Not a valid date"
        dlpDateRange(x%) = ""
        dlpDateRange(x%).SetFocus
        Exit Function
    End If
 End If
Next x%

For x% = 0 To 1
 If Len(medDOW(x%)) > 0 Then
    If medDOW(x%) < 0 Or medDOW(x%) > 7 Then
        MsgBox "Day of Week must be between 1 and 7"
        medDOW(x%).SetFocus
        Exit Function
    End If
  End If
Next x%

If Len(txtIncident) > 0 Then
    If txtIncident <> "Y" And txtIncident <> "N" Then
        MsgBox "Incident must be Y/N or blank"
        txtIncident.SetFocus
        Exit Function
    End If
End If

If Len(txtIncentive) > 0 Then
    If txtIncentive <> "Y" And txtIncentive <> "N" Then
        If glbCompSerial = "S/N - 2214W" Then 'Casey House  - Ticket #15276
            MsgBox "HOOPP must be Y/N or blank"
        Else
            MsgBox "Incentive must be Y/N or blank"
        End If
        txtIncentive.SetFocus
        Exit Function
    End If
End If

If Len(txtSeniority) > 0 Then
    If UCase(txtSeniority) <> "Y" And UCase(txtSeniority) <> "N" Then
        MsgBox "Seniority must be Y/N or blank"
        txtSeniority.SetFocus
        Exit Function
    End If
End If

If Not clpPT.ListChecker Then
'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Len(txtHours) > 0 Then
    If Not IsNumeric(txtHours) Then
        MsgBox "Hours must be number"
        txtHours.SetFocus
        Exit Function
    End If
End If

For x% = 0 To 1
    If Len(txtPoint(x%)) > 0 Then
        If Not IsNumeric(txtPoint(x%)) Then
            MsgBox "Point must be number"
            txtPoint(x%).SetFocus
            Exit Function
        End If
    End If
Next x%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then 'Hemu - 05/14/2003 Check for the blank - Begin
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If 'Hemu - 05/14/2003 Check for the blank - End

If Len(txtPoint(0)) > 0 And Len(txtPoint(1)) > 0 Then
    If Val(txtPoint(0)) > Val(txtPoint(1)) Then
            MsgBox "From Point can not be greater than To Point"
            txtPoint(0).SetFocus
            Exit Function
    End If
End If

If glbCompSerial = "S/N - 2214W" Then
    If Len(clpChrgCode) > 0 Then
        If clpChrgCode.Caption = "Unassigned" Then
            MsgBox "If Attendance-Department is entered it must be valid"
            clpChrgCode.SetFocus
            Exit Function
        End If
    End If
    If Len(clpCode(3)) > 0 Then
        If clpCode(3).Caption = "Unassigned" Then
            MsgBox "Invalid Attendance-Fund code"
            clpCode(3).SetFocus
            Exit Function
        End If
    End If
End If

If glbCompSerial = "S/N - 2396W" Then  'Oshawa CHC - Ticket #17323
    If Len(clpChrgCode) > 0 Then
        If clpChrgCode.Caption = "Unassigned" Then
            MsgBox lStr("If G/L # is entered it must be valid")
            clpChrgCode.SetFocus
            Exit Function
        End If
    End If
End If
CriCheck = True
End Function

Private Sub optGrouping_Click(Index As Integer, Value As Integer)
    'Release 8.0 - Ticket #22682: Hide Name
    If Index = 2 Then
        chkHideName.Visible = True
    Else
        chkHideName.Visible = False
    End If
End Sub

Private Sub txtHours_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtIncentive_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtIncentive_KeyUp(KeyCode As Integer, Shift As Integer)
txtIncentive.Text = UCase(txtIncentive.Text)
End Sub

Private Sub txtIncident_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtIncident_KeyUp(KeyCode As Integer, Shift As Integer)
txtIncident.Text = UCase(txtIncident.Text)
End Sub

Private Sub txtPoint_Change(Index As Integer)
If glbBurlTech Then
If Index = 0 Then
    If Len(txtPoint(0).Text) > 0 Then
        chkPointType1.Visible = True
        chkPointType2.Visible = True
    Else
        chkPointType1.Visible = False
        chkPointType2.Visible = False
    End If
End If
End If
End Sub

Private Sub txtSeniority_GotFocus()
Call SetPanHelp(Me.ActiveControl)
MDIMain.panHelp(0).Caption = Replace(MDIMain.panHelp(0).Caption, "s", "S")
End Sub

Private Sub txtSeniority_KeyUp(KeyCode As Integer, Shift As Integer)
txtSeniority.Text = UCase(txtSeniority.Text)
End Sub

Private Sub txtShift_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Cri_Job()
Dim EECri As String

If Len(clpJOB.Text) > 0 Then
    EECri = "{HRATTWRK.AD_JOB} = '" & clpJOB.Text & "' "
    HisSQL = HisSQL & " AND (HR_ATTENDANCE.AD_JOB = '" & clpJOB.Text & "') "
    HisSQL1 = HisSQL1 & " AND (HR_ATTENDANCE_HISTORY.AH_JOB = '" & clpJOB.Text & "') "
        
    'Ticket #29107 - All selection criteria except From / To Date Range
    HisSQLNoDate = HisSQLNoDate & " AND (HR_ATTENDANCE.AD_JOB = '" & clpJOB.Text & "') "
    HisSQL1NoDate = HisSQL1NoDate & " AND (HR_ATTENDANCE_HISTORY.AH_JOB = '" & clpJOB.Text & "') "
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

Private Sub txtChargeCode_GotFocus()
   Call SetPanHelp(Me.ActiveControl)     'laura 03/03/98
End Sub

Private Sub SELATTWRK()
Dim xlen, xxx, xx1
Dim SQLQ, xNum, xNumLE, xNumBF, I, xPoint, xNumnbr
Dim rsAttEmp As New ADODB.Recordset
Dim xFieldList
Dim rsAttWRK As New ADODB.Recordset
On Error GoTo AttWrkError
'AD_WHRS AS Total Points of employees - Special
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
Set CN = New ADODB.Connection
CN.CommandTimeout = 600
CN.Open glbAdoIHRDB
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
CN.BeginTrans
CN.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
CN.CommitTrans

MDIMain.panHelp(0).FloodPercent = 30

xFieldList = Get_Fields(CN, "HR_ATTENDANCE", "AD_ATT_ID,AD_DHRS,AD_WHRS,AD_DISCIPLINE_TABL,AD_DISCIPLINE")

SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
SQLQ = SQLQ & in_SQL(glbIHRDBW)

SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE "
SQLQ = SQLQ & " WHERE " & HisSQL

If Val(txtPoint(0)) <> 0 Or Val(txtPoint(1)) <> 0 Then
    If glbBurlTech And chkPointType2.Visible And chkPointType2.Value Then
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE GROUP BY AD_EMPNBR "
        SQLQ = SQLQ & " HAVING ( SUM (AD_LEPOINT) >= " & Val(txtPoint(0))
        If Len(txtPoint(1)) > 0 Then SQLQ = SQLQ & " AND  SUM(AD_LEPOINT) <= " & Val(txtPoint(1))
        SQLQ = SQLQ & "))"
    Else
        SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT AD_EMPNBR FROM HR_ATTENDANCE GROUP BY AD_EMPNBR "
        SQLQ = SQLQ & " HAVING ( SUM (AD_POINT) >= " & Val(txtPoint(0))
        If Len(txtPoint(1)) > 0 Then SQLQ = SQLQ & " AND  SUM(AD_POINT) <= " & Val(txtPoint(1))
        SQLQ = SQLQ & "))"
    End If
End If


MDIMain.panHelp(0).FloodPercent = 45
CN.BeginTrans
CN.Execute SQLQ
CN.CommitTrans

MDIMain.panHelp(0).FloodPercent = 60
If chkInclAtt.Value Then
    SQLQ = "INSERT INTO HRATTWRK (" & xFieldList & ",AD_WRKEMP) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & " SELECT " & Replace(xFieldList, "AD_", "AH_") & ",'" & glbUserID & "' AS AD_WRKEMP "
    SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
    SQLQ = SQLQ & " WHERE " & HisSQL1
    If Val(txtPoint(0)) <> 0 Or Val(txtPoint(1)) <> 0 Then
        If glbBurlTech And chkPointType2.Visible And chkPointType2.Value Then
            SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY GROUP BY AH_EMPNBR "
            SQLQ = SQLQ & " HAVING ( SUM (AH_LEPOINT) >= " & Val(txtPoint(0))
            If Len(txtPoint(1)) > 0 Then SQLQ = SQLQ & " AND  SUM(AH_LEPOINT) <= " & Val(txtPoint(1))
            SQLQ = SQLQ & "))"
        Else
            SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT AH_EMPNBR FROM HR_ATTENDANCE_HISTORY GROUP BY AH_EMPNBR "
            SQLQ = SQLQ & " HAVING ( SUM (AH_POINT) >= " & Val(txtPoint(0))
            If Len(txtPoint(1)) > 0 Then SQLQ = SQLQ & " AND  SUM(AH_POINT) <= " & Val(txtPoint(1))
            SQLQ = SQLQ & "))"
        End If
    End If
    MDIMain.panHelp(0).FloodPercent = 75
    CN.BeginTrans
    CN.Execute SQLQ
    CN.CommitTrans
End If

SQLQ = "UPDATE HRATTWRK SET AD_POINT = 0 WHERE AD_POINT IS NULL "
gdbAdoIhr001W.Execute SQLQ
SQLQ = "UPDATE HRATTWRK SET AD_LEPOINT = 0 WHERE AD_LEPOINT IS NULL"
gdbAdoIhr001W.Execute SQLQ
    
'Bring forward the points
'AD_WHRS = BF Points for each employee
'AD_DHRS = AD_WHRS/records count for each employee
If Len(dlpDateRange(0)) > 0 And chkBF Then
    Call Pause(1)
    SQLQ = "SELECT AD_EMPNBR AS EMPNO, count(AD_EMPNBR) AS EMPCOUNT FROM HRATTWRK "
    SQLQ = SQLQ & "WHERE AD_WRKEMP='" & glbUserID & "' " 'Ticket #11710
    SQLQ = SQLQ & "GROUP BY AD_EMPNBR "
    rsAttWRK.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockReadOnly
    xxx = 0: I = 0
    If Not rsAttWRK.EOF Then
        xxx = rsAttWRK.RecordCount
    End If
    MDIMain.panHelp(0).FloodPercent = 0
    Do While Not rsAttWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (I / xxx) * 100
        xNum = 0: xNumLE = 0
        xNumnbr = rsAttWRK("EMPNO")
        
        'Sum the B/F Points in Attendance
        SQLQ = "SELECT SUM(AD_POINT) AS SUMPOINT "
        If glbBurlTech Then
            SQLQ = SQLQ & ",SUM(AD_LEPOINT) AS SUMLEPOINT "
        End If
        SQLQ = SQLQ & " FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xNumnbr & " "
        SQLQ = SQLQ & " AND AD_DOA < " & Date_SQL(dlpDateRange(0)) & " "
        
        'Ticket #29107 - All selection criteria except From / To Date Range
        If Len(HisSQLNoDate) > 0 Then
            SQLQ = SQLQ & " AND " & HisSQLNoDate
        End If
        
        SQLQ = SQLQ & "GROUP BY AD_EMPNBR "
        rsAttEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsAttEmp.EOF Then
            If Not IsNull(rsAttEmp("SUMPOINT")) Then
                xNum = xNum + rsAttEmp("SUMPOINT")
            End If
            If glbBurlTech Then
                If Not IsNull(rsAttEmp("SUMLEPOINT")) Then
                    xNumLE = xNumLE + rsAttEmp("SUMLEPOINT")
                End If
            End If
        End If
        rsAttEmp.Close
        
        'Sum the B/F points in Attendance History
        If chkInclAtt.Value Then
            SQLQ = "SELECT SUM(AH_POINT) AS SUMPOINT "
            If glbBurlTech Then
                SQLQ = SQLQ & ",SUM(AH_POINT) AS SUMLEPOINT "
            End If
            SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & xNumnbr & " "
            SQLQ = SQLQ & " AND AH_DOA < " & Date_SQL(dlpDateRange(0)) & " "
            
            'Ticket #29107 - All selection criteria except From / To Date Range
            If Len(HisSQL1NoDate) > 0 Then
                SQLQ = SQLQ & " AND " & HisSQL1NoDate
            End If
            
            SQLQ = SQLQ & "GROUP BY AH_EMPNBR "
            rsAttEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsAttEmp.EOF Then
                If Not IsNull(rsAttEmp("SUMPOINT")) Then
                    xNum = xNum + rsAttEmp("SUMPOINT")
                End If
                If glbBurlTech Then
                    If Not IsNull(rsAttEmp("SUMLEPOINT")) Then
                        xNumLE = xNumLE + rsAttEmp("SUMLEPOINT")
                    End If
                End If
            End If
            rsAttEmp.Close
        End If
        
        'Update the B/F Points
        SQLQ = "UPDATE HRATTWRK SET AD_WHRS = " & xNum & " "
        'Ticket #12101, Franks Nov 17, 2006, Get Total Point = BF Points + Points in this date
        SQLQ = SQLQ & ",AD_DHRS = " & xNum / rsAttWRK("EMPCOUNT") & " "
        If glbBurlTech Then
            SQLQ = SQLQ & ",AD_MACHINE_RATE = " & xNumLE & " "
        End If
        SQLQ = SQLQ & "WHERE AD_EMPNBR = " & xNumnbr & " "
        gdbAdoIhr001W.Execute SQLQ

        rsAttWRK.MoveNext
    Loop
    rsAttWRK.Close
    
    Call Pause(5)
Else
    'Clear any B/F Points because user did not check B/F Points
    SQLQ = "UPDATE HRATTWRK SET AD_WHRS = 0,AD_DHRS = 0 "
    If glbBurlTech Then
        SQLQ = SQLQ & ",AD_MACHINE_RATE = 0 "
    End If
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute SQLQ
    gdbAdoIhr001W.CommitTrans
    Call Pause(5)
End If

SQLQ = ""
HisSQL = ""
HisSQL1 = ""

'Ticket #29107 - All selection criteria except From / To Date Range
HisSQLNoDate = ""
HisSQL1NoDate = ""

CN.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

AttWrkError:
    CN.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    If False Then
        Resume   ' for debugging
    End If
    Exit Sub
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

