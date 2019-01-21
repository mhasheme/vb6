VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRTrainMatrix 
   Caption         =   "Training Matrix Report"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9510
   ScaleWidth      =   11745
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkExclCONP 
      Alignment       =   1  'Right Justify
      Caption         =   "Exclude Employment Status of CONP"
      Height          =   285
      Left            =   120
      TabIndex        =   59
      Tag             =   "Check to Exclude Employees with CONP Employment Status"
      Top             =   7560
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CheckBox chkExclRET 
      Alignment       =   1  'Right Justify
      Caption         =   "Exclude Employment Status of RET"
      Height          =   285
      Left            =   120
      TabIndex        =   58
      Tag             =   "Check to Exclude Employees with RET Employment Status"
      Top             =   7860
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.CheckBox chkShowAllEmp 
      Alignment       =   1  'Right Justify
      Caption         =   "Show All Employees With or Without Training Records "
      Height          =   225
      Left            =   120
      TabIndex        =   20
      Tag             =   "If checked - Show All Employees"
      Top             =   7200
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.CheckBox chkShowEmp 
      Alignment       =   1  'Right Justify
      Caption         =   "Show All Employees With any Training Record"
      Height          =   225
      Left            =   120
      TabIndex        =   19
      Tag             =   "If checked -Show All Employees"
      Top             =   6480
      Width           =   4755
   End
   Begin VB.CheckBox chkTrainingList 
      Alignment       =   1  'Right Justify
      Caption         =   "Use Training Plan to generate Training Matrix Report"
      Height          =   225
      Left            =   120
      TabIndex        =   57
      Top             =   8520
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Frame fraGroup 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   6000
      TabIndex        =   46
      Top             =   8400
      Visible         =   0   'False
      Width           =   4335
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Tag             =   "Final sorting of records - no totals"
         Top             =   1275
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Tag             =   "Third level of grouping records"
         Top             =   960
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Tag             =   "Second level of grouping records"
         Top             =   645
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   1860
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Tag             =   "First Level of grouping records"
         Top             =   330
         Width           =   2325
      End
      Begin VB.Label Label5 
         Caption         =   "for Training Plan"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   2280
         TabIndex        =   56
         Top             =   0
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Final Sort"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   55
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #3"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   990
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #2"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   53
         Top             =   675
         Width           =   885
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Grouping #1"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   52
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblReportGrp 
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
         TabIndex        =   51
         Top             =   0
         Width           =   1575
      End
   End
   Begin VB.CheckBox chkRenewalExce 
      Alignment       =   1  'Right Justify
      Caption         =   "Any course with a renewal date less than today’s date"
      Height          =   225
      Left            =   120
      TabIndex        =   45
      Top             =   6840
      Visible         =   0   'False
      Width           =   4740
   End
   Begin VB.CheckBox chkCoursesTaken 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Courses Taken Only "
      Height          =   225
      Left            =   120
      TabIndex        =   25
      Top             =   8160
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2220
   End
   Begin INFOHR_Controls.CodeLookup clpCrsType 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   4365
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowDescription =   0   'False
      TABLName        =   "ESCT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.CheckBox chkReqCourses 
      Alignment       =   1  'Right Justify
      Caption         =   "Required Courses Only   "
      Height          =   225
      Left            =   120
      TabIndex        =   23
      Top             =   6120
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CheckBox chkLegisCourses 
      Alignment       =   1  'Right Justify
      Caption         =   "And Legislated Courses Only"
      Height          =   225
      Left            =   2460
      TabIndex        =   24
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2115
      MaxLength       =   4
      TabIndex        =   16
      Tag             =   "00-Employee Position Shift"
      Top             =   5036
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1695
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
      Left            =   1800
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2025
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
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1365
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
      Top             =   1035
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   12
      Tag             =   "00-Enter Administered By Code"
      Top             =   3705
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   13
      Tag             =   "00-Enter Section Code"
      Top             =   4035
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   11
      Tag             =   "00-Enter Region Code"
      Top             =   3360
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2370
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3450
      TabIndex        =   8
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2700
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   7
      Tag             =   "40-Date from and including this date forward"
      Top             =   2700
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Enter Position Code"
      Top             =   3030
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpPosGroup 
      Height          =   285
      Left            =   7380
      TabIndex        =   10
      Tag             =   "00-Position Group  Code"
      Top             =   3030
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10560
      Top             =   8160
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
   Begin Threed.SSCheck chkShowEmp1 
      Height          =   225
      Left            =   6120
      TabIndex        =   21
      Tag             =   "If X-Show All Employees"
      Top             =   8640
      Visible         =   0   'False
      Width           =   3540
      _Version        =   65536
      _ExtentX        =   6244
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Show All Employees With any Training Record"
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
      Alignment       =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCrsCode 
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   4695
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowDescription =   0   'False
      TABLName        =   "ESCD"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpProv 
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Tag             =   "31-Province of Residence - Code"
      Top             =   5380
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.CodeLookup clpProvEmp 
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Tag             =   "31-Province of Employment - Code"
      Top             =   5720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin Threed.SSCheck chkShowAllEmp1 
      Height          =   225
      Left            =   6000
      TabIndex        =   22
      Tag             =   "If X-Show All Employees"
      Top             =   8880
      Visible         =   0   'False
      Width           =   4275
      _Version        =   65536
      _ExtentX        =   7541
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Show All Employees With or Without Training Records "
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
      Alignment       =   1
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. of Residence"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   44
      Top             =   5425
      Width           =   1365
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Prov. of Employment"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   43
      Top             =   5765
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6240
      TabIndex        =   42
      Top             =   3075
      Width           =   1035
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   41
      Top             =   3077
      Width           =   975
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   40
      Top             =   4413
      Width           =   900
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Course Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   39
      Top             =   4747
      Width           =   915
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shift"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   38
      Top             =   5081
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   37
      Top             =   2075
      Width           =   630
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   36
      Top             =   405
      Width           =   555
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   35
      Top             =   739
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   34
      Top             =   1407
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   33
      Top             =   1741
      Width           =   450
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   32
      Top             =   2409
      Width           =   1290
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
      TabIndex        =   31
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   2743
      Width           =   1095
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   1073
      Width           =   615
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3411
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   3745
      Width           =   1125
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   4079
      Width           =   540
   End
End
Attribute VB_Name = "frmRTrainMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CodeCodes(10, 2)

Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc, xKeyPress, xTxtDiv, xLblDivDesc
Dim LastlastID, LastlastNme, LastFirstNme, xTxtEEID, xLblEEName
Dim PosFlag As Boolean, strShift, strPosCode, strPosGrp
Dim strReqCourses As String
Dim xTrainMatrixPath
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
 
Private Sub XLSwriter_Bur()
Dim SQLQ As String
Dim exApp As Excel.Application, exBook As Excel.Workbook, exSheet As Excel.Worksheet
Dim xlsFileTmp As String, xlsFileMat As String
Dim rsIn As New ADODB.Recordset, rsOUT As New ADODB.Recordset, rsEMP As New ADODB.Recordset, rsJOB As New ADODB.Recordset, rsCurJob As New ADODB.Recordset
Dim xRow As Long, xCol As Long, xwCol As Long
Dim xType As String, strTemp As String, strDate As String
Dim NewDateFormat As String, flgReqC As Boolean, strDisp As String, xMax As Long, retval As Long
Dim xStartDate As String

On Error GoTo Err_XLS

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainJobTmp.xls"
    'Ticket# 8293
    If glbLinamar Then
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Reports\TrainJobMat" & Trim(glbUserID) & ".xls"
    Else
        xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "TrainJobMat" & Trim(glbUserID) & ".xls"
    End If

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
        'Create new WorkBook of Excel
    Set exApp = New Excel.Application
    Set exBook = exApp.Workbooks.Open(xlsFileMat)

    
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    
    SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_JOB, HRJOB.JB_DESCR FROM HR_JOB_HISTORY INNER JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "WHERE (1 = 1) and JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & getWSQLQ(False) & ")"
    If clpJob <> "" Then SQLQ = SQLQ & " AND JH_JOB = '" & clpJob.Text & "' "
    SQLQ = SQLQ & " ORDER BY JB_DESCR ASC"
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsJOB.EOF = False And rsJOB.BOF = False Then
        Do
            
            'Courses
            'If glbOracle Or glbSQL Then 'Oracle 9i and SQL
                SQLQ = "SELECT DISTINCT HR_JOB_COURSE.PC_CRSCODE as ES_CRSCODE, HRTABL.TB_DESC AS ES_CRSDESC, tblType.TB_KEY AS ES_CRSTYPE, tblTYPE.TB_DESC as ES_CTYPEDESC, HR_JOB_COURSE.PC_JOB as ES_JOB, HR_JOB_COURSE.PC_LEGISLATED as ES_LEGISLATE "
                SQLQ = SQLQ & "FROM HR_JOB_COURSE INNER JOIN HRTABL ON HR_JOB_COURSE.PC_CRSCODE = HRTABL.TB_KEY AND HR_JOB_COURSE.PC_CRSCODE_TABL = HRTABL.TB_NAME LEFT OUTER JOIN HRTABL tblType ON HRTABL.TB_USR1 = tblType.TB_KEY "
            'End If
            SQLQ = SQLQ & "WHERE 1=1 AND PC_JOB = '" & rsJOB("JH_JOB") & "' and tblType.TB_NAME='ESCT' "
            If clpCrsCode <> "" Then SQLQ = SQLQ & " AND PC_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
            If chkLegisCourses Then SQLQ = SQLQ & " AND PC_LEGISLATED <> 0"
            SQLQ = SQLQ & " ORDER BY tblType.TB_KEY ASC"
            rsIn.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
            
            'Employee Info
            SQLQ = "SELECT DISTINCT ED_EMPNBR, ED_SURNAME, ED_FNAME, ED_DOH FROM HREMP WHERE "
            SQLQ = SQLQ & " ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB='" & rsJOB("JH_JOB") & "') AND "
            SQLQ = SQLQ & getWSQLQ(False)
            
            rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            'Hemu - Since if there are more than 260 courses in the table it gives the Subscript out of range error
            '       and also Excel cannot take more than 124cols, the user may select only the courses
            '       needed in the report, made the following change to only select from the HRTABL the
            '       courses based on the selection criteria given by the user
            
            If rsIn.EOF = False And rsIn.BOF = False And rsEMP.EOF = False And rsEMP.BOF = False Then
                xMax = rsEMP.RecordCount
                If xMax > 250 Then
                    retval = MsgBox("This Training Matrix exceeds 250 employees, Click Yes to Continue, No to refine the query", vbYesNo + vbQuestion, "Columns Exceeded")
                    If retval = vbNo Then
                        GoTo exH
                    End If
                End If
                
                exBook.Worksheets(1).Copy After:=exBook.Worksheets(exBook.Worksheets.count)
                Set exSheet = exBook.Worksheets(exBook.Worksheets.count)
                
                exSheet.name = Left(rsJOB("JB_DESCR"), 30)
                  'exSheet.Cells(8, 1) = "Position: " & rsJob("JB_DESCR")
                
                exSheet.Cells(2, 9) = "Date: " & Format(Now, "mmm dd, yyyy")
                exSheet.Cells(3, 9) = "Time: " & Time$
                If Len(clpDept) > 0 Then
                    exSheet.Cells(2, 1) = lStr("Department: ") & clpDept.Caption
                End If
                If Len(clpJob.Text) > 0 Then
                    exSheet.Cells(3, 1) = "Position: " & rsJOB("JB_DESCR") 'rsJob("JH_JOB")
                End If
                 
'                If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
'                    exSheet.Cells(4, 5) = "No date entered"
'                Else
'                    strTemp = ""
'                    If IsDate(dlpDateRange(0).Text) Then
'                        strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
'                    End If
'                    If IsDate(dlpDateRange(1).Text) Then
'                        strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
'                    End If
'                    exSheet.Cells(4, 5) = strTemp
'                End If
        
                xType = ""
                xRow = 12
                Do
                    xCol = 3
                    'Course Type
                    If rsIn("ES_CRSTYPE") <> xType Then
                        exSheet.Cells(xRow, 1) = rsIn("ES_CTYPEDESC")
                        exSheet.Rows(xRow).EntireRow.Interior.Color = RGB(210, 210, 210)
                        exSheet.Cells(xRow, 1).Font.Bold = True
                        exSheet.Cells(xRow, 1).HorizontalAlignment = xlLeft
                        xType = rsIn("ES_CRSTYPE")
                        xRow = xRow + 1
                    ElseIf IsNull(rsIn("ES_CRSTYPE")) And xType = "" Then
                        xType = "None"
                        exSheet.Cells(xRow, 1) = xType
                        exSheet.Rows(xRow).EntireRow.Interior.Color = RGB(210, 210, 210)
                        exSheet.Cells(xRow, 1).HorizontalAlignment = xlLeft
                        exSheet.Cells(xRow, 1).Font.Bold = True
                        xRow = xRow + 1
                    End If
                    'Course Description
                    exSheet.Cells(xRow, 1) = rsIn("ES_CRSDESC")
        
                    rsEMP.MoveFirst
                    Do
                        'Get the Position Start Date - Ticket #13158
                        SQLQ = "SELECT JH_EMPNBR, JH_SDATE FROM HR_JOB_HISTORY"
                        SQLQ = SQLQ & " WHERE JH_EMPNBR = " & rsEMP("ED_EMPNBR")
                        SQLQ = SQLQ & " AND JH_CURRENT<>0 AND JH_JOB='" & rsJOB("JH_JOB") & "' "
                        rsCurJob.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
                        If Not rsCurJob.EOF Then
                            xStartDate = rsCurJob("JH_SDATE")
                        Else
                            xStartDate = ""
                        End If
                        rsCurJob.Close
                    
                        exSheet.Cells(10, xCol) = rsEMP("ED_SURNAME") & ", " & rsEMP("ED_FNAME")
                        'exSheet.Cells(11, xCol) = rsEMP("ED_EMPNBR") & "  " & rsEMP("ED_DOH")
                        exSheet.Cells(11, xCol) = rsEMP("ED_EMPNBR") & "  " & xStartDate
                        
                        SQLQ = "SELECT ES_CRSCODE, ES_DATCOMP, ES_RENEW, ES_RESULTS FROM HREDSEM WHERE "
                        SQLQ = SQLQ & "ES_EMPNBR=" & rsEMP("ED_EMPNBR") & " AND ES_CRSCODE='" & rsIn("ES_CRSCODE") & "' "
                        If rsIn("ES_CRSTYPE") <> "" Then
                            SQLQ = SQLQ & " AND ES_CTYPE='" & rsIn("ES_CRSTYPE") & "'"
                        End If
                        SQLQ = SQLQ & " ORDER BY ES_DATCOMP DESC"
                        rsOUT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
                        If rsOUT.EOF = False And rsOUT.BOF = False Then
                  
                            If Not IsNull(rsOUT("ES_CRSCODE")) Then
                                'Green:   Training in Good Standing
                                'Required Course with no renewal date
                                'Required Course with a renewal date greater than 30 days from today
                                'A non-required course with no renewal date
                                '
                                'Yellow:    Training will Expire in Thirty Days
                                'Any course with a renewal date within 30 days of today
                                '
                                'Red:   Training Not in Good Standing
                                'Required Course with no completed date.
                                'Required Course not taken (No Continuing Education record)
                                'Any course with a renewal date less than today’s date
                                    
                                'Check If Required Courses
                                If IsDate(rsOUT("ES_DATCOMP")) Then
                                    strDate = Format(rsOUT("ES_DATCOMP"), NewDateFormat)
                                    strDisp = "Good"
                                Else
                                    strDate = ""
                                    strDisp = ""
                                End If
                                
                                If Len(rsOUT("ES_RESULTS")) > 0 Then
                                    Select Case rsOUT("ES_RESULTS")
                                    Case "N/A", "N\A", "MR", "PMR"
                                        strTemp = rsOUT("ES_RESULTS")
                                        strDisp = "Good"
                                    Case Else
                                        strTemp = strDate
                                    End Select
                                Else
                                    strTemp = strDate
                                End If
                                
                                exSheet.Cells(xRow, xCol) = strTemp
                                
    '                            If rsIn("ES_REQUIRED") = "R" Then
                                    flgReqC = True
    '                            Else
    '                                flgReqC = False
    '                            End If
                                'for Good Standing - Green
                                If flgReqC And Not IsDate(rsOUT("ES_RENEW")) Then
                                        strDisp = "Good"
                                End If
                                If flgReqC And IsDate(rsOUT("ES_RENEW")) Then
                                    If DateDiff("d", Now, CVDate(rsOUT("ES_RENEW"))) > 30 Then
                                        strDisp = "Good"
                                    End If
                                End If
                                If Not flgReqC And Not IsDate(rsOUT("ES_RENEW")) And IsDate(rsOUT("ES_DATCOMP")) Then
                                    strDisp = "Good"
                                End If
                                'Yellow:    Training will Expire in Thirty Days
                                If IsDate(rsOUT("ES_RENEW")) Then
                                    If DateDiff("d", Now, CVDate(rsOUT("ES_RENEW"))) >= 0 And DateDiff("d", Now, CVDate(rsOUT("ES_RENEW"))) <= 30 Then
                                        strDisp = "Expire"
                                    End If
                                End If
                                'Red:   Training Not in Good Standing
                                If IsDate(rsOUT("ES_RENEW")) Then
                                    If DateDiff("d", Now, CVDate(rsOUT("ES_RENEW"))) < 0 Then
                                        strDisp = "Not Good"
                                    End If
                                End If
                                If flgReqC And Not IsDate(rsOUT("ES_DATCOMP")) Then
                                    strDisp = "Not Good"
                                End If
                                
                                
                                'Hemu
    '                            Select Case strDisp
    '                            Case "Good" 'green
    '                                exSheet.Cells(xRow, xCol).Interior.ColorIndex = 50
    '                            Case "Expire" ' red
    '                                exSheet.Cells(xRow, xCol).Interior.ColorIndex = 36
    '                            Case "Not Good" 'yellow
    '                                exSheet.Cells(xRow, xCol).Interior.ColorIndex = 3
    '                            End Select
                                'Hemu
                            End If
                        End If
                        rsOUT.Close
                        rsEMP.MoveNext
                        xCol = xCol + 1
                        If xCol - 3 > xMax Then
                            Exit Do
                        End If
                    Loop Until rsEMP.EOF
                    rsIn.MoveNext
                    xRow = xRow + 1
                Loop Until rsIn.EOF
                
                If xCol < 11 Then
                    exSheet.Range(exSheet.Cells(12, 1), exSheet.Cells(xRow, 11)).Borders.LineStyle = xlThin
                    exSheet.Range(exSheet.Cells(12, 1), exSheet.Cells(xRow, 11)).Borders.LineStyle = xlSolid
                    exSheet.Range(exSheet.Cells(10, 1), exSheet.Cells(10, 11)).Columns.AutoFit
                Else
                    exSheet.Range(exSheet.Cells(12, 1), exSheet.Cells(xRow, xCol)).Borders.LineStyle = xlThin
                    exSheet.Range(exSheet.Cells(12, 1), exSheet.Cells(xRow, xCol)).Borders.LineStyle = xlSolid
                    exSheet.Range(exSheet.Cells(10, 1), exSheet.Cells(10, xCol)).Columns.AutoFit
                End If
                
                xRow = xRow + 3
                exSheet.Cells(xRow, 1) = "Manager:____________________________________"
                exSheet.Cells(xRow, 5) = "Date:__________________"
                
'                If xRow < 34 Then
'                    xRow = 34
'                    exSheet.Cells(xRow, 1) = "Date Issued: " & Format(Now, "mmm dd, yyyy")
'                    xRow = 35
'                    exSheet.Cells(xRow, 1) = "Revision Date: "
'                Else
'                    xRow = xRow + 2
'                    exSheet.Cells(xRow, 1) = "Date Issued: " & Format(Now, "mmm dd, yyyy")
'                    xRow = xRow + 1
'                    exSheet.Cells(xRow, 1) = "Revision Date: "
'                End If
'                exSheet.PageSetup.RightFooter = "Form 1808-1"
                
                
                If xCol < 11 Then
                    exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(1, 11)).Merge
                    exSheet.Range(exSheet.Cells(9, 2), exSheet.Cells(9, 11)).Merge
                    exSheet.PageSetup.PrintArea = exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(xRow, 11)).AddressLocal
                Else
                    exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(1, xCol)).Merge
                    exSheet.Range(exSheet.Cells(9, 2), exSheet.Cells(9, xCol)).Merge
                    exSheet.PageSetup.PrintArea = exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(xRow, xCol)).AddressLocal
                End If
                
                exSheet.PageSetup.Orientation = xlLandscape
                exSheet.PageSetup.Zoom = False
                exSheet.PageSetup.FitToPagesTall = 1
                exSheet.PageSetup.FitToPagesWide = 1
                exSheet.Range("A1").Select
                'Save new Excel file as XLS
                'exBook.SaveAs "C:\TrainMat.xls"
                exBook.Save
                Set exSheet = Nothing
            End If
            rsEMP.Close
            rsIn.Close
            rsJOB.MoveNext
        Loop Until rsJOB.EOF
        Set exSheet = exBook.Worksheets(1)
        exApp.DisplayAlerts = False
        exSheet.Delete
        exApp.DisplayAlerts = True
        exBook.Worksheets(1).Activate
        exBook.Save
        Set exBook = Nothing
        exApp.Visible = True
        Set exApp = Nothing
    Else
        MsgBox "There are no Records matching this criteria", vbInformation + vbOKOnly, "No Records"
        exBook.Close False
        exApp.Quit
        Set exApp = Nothing
    End If
    
    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
'    If Not LanchXlsW98(xlsFileMat) Then
'        Shell "cmd /c " & GetShortName(xlsFileMat)
'    End If
exH:
    
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
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Err = 76 Then
        MsgBox Err.Description & " to save the Training Matrix Report." & vbCrLf & "Please check Company Preference under Setup Menu."
        Exit Sub
    End If
    If Not exApp Is Nothing Then
        If exApp.Visible = False Then
            exApp.Quit
        End If
        Set exApp = Nothing
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriter", "", "Select")
'Resume Next
End Sub

Private Sub chkExclCONP_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkExclRET_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkReqCourses_Click()
'Hemu - 08/20/2003 Begin - Ticket 4609
    If chkReqCourses Then
        chkLegisCourses.Visible = True
    Else
        chkLegisCourses.Visible = False
    End If
'Hemu - 08/20/2003 End
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    MDIMain.MainToolBar.ButtonS("preview").Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    If glbFormCaption = "Training Matrix Report" Then
    '    cmdView.Enabled = False
        Screen.MousePointer = HOURGLASS
        x% = Cri_SetAll()
        Screen.MousePointer = DEFAULT
        MDIMain.Timer1.Enabled = True
    '    cmdView.Enabled = True
    End If
    
    'Ticket #21709 Franks 03/12/2012
    If glbFormCaption = "Training Plan Report" Then
        Screen.MousePointer = HOURGLASS
        Call set_PrintState(False)
        x% = Cri_SetAl_TrainPlan
        Me.vbxCrystal.Destination = 0
        MDIMain.Timer1.Enabled = False
        Screen.MousePointer = DEFAULT
        Me.vbxCrystal.Action = 1
        vbxCrystal.Reset
        MDIMain.Timer1.Enabled = True
        Call set_PrintState(True)

    End If
MDIMain.MainToolBar.ButtonS("preview").Enabled = True
End If
Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub chkShowAllEmp_Click(Value As Integer)
    'If chkShowAllEmp.Visible Then 'Ticket #28174 Franks 05/10/2016 for WFC
    '    If chkShowAllEmp.Value Then
    '        If chkShowEmp.Value Then
    '            chkShowEmp.Value = False
    '        End If
    '    End If
    'End If
'End Sub

'Private Sub chkShowEmp_Click(Value As Integer)
    'If chkShowAllEmp.Visible Then 'Ticket #28174 Franks 05/10/2016 for WFC
    '    If chkShowEmp.Value Then
    '        If chkShowAllEmp.Value Then
    '            chkShowAllEmp.Value = False
    '        End If
    '    End If
    'End If
'End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
MDIMain.MainToolBar.ButtonS("print").Enabled = False
MDIMain.mnu_F_Print.Enabled = False
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Me.Caption = glbFormCaption
 
Screen.MousePointer = HOURGLASS

If glbWFC Then 'Ticket #25911 Franks 01/27/2015
    clpJob.TextBoxWidth = 1265
End If

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

'Hemu - 08/20/2003 Begin - Ticket 4609
'If glbCompSerial = "S/N - 2288W" Then
    chkReqCourses.Visible = True
'Else
'    chkReqCourses.Visible = False
'End If
'Hemu - 08/20/2003 End

'Ticket #22274 - City of Chatham-Kent - Option to use Training Plan instead of Continuing Education
If glbCompSerial = "S/N - 2188W" Then
    chkTrainingList.Visible = True
Else
    chkTrainingList.Visible = False
End If

If gsTRAININGMATRIX Then
    xTrainMatrixPath = GetComPreferEmail("TRAININGMATRIX")
End If
If Len(xTrainMatrixPath) = 0 Then
    xTrainMatrixPath = glbIHRREPORTS
End If

If glbWFC Then
    chkReqCourses.Visible = False
    chkCoursesTaken.Enabled = False 'Ticket #21499 Franks 01/26/2012
    'chkCoursesTaken.Visible = True
    chkCoursesTaken.Visible = False
    chkRenewalExce.Visible = True 'Ticket #20479 Franks 07/11/2011
End If

Call setRptCaption(Me)

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbWFC And glbFormCaption = "Training Matrix Report" Then
    lblFromTo.FontBold = True
    lblSection.FontBold = True
    If Len(glbPlantCode) > 0 Then
        clpCode(5).Text = glbPlantCode
        'Ticket #20947 Franks 09/13/2011
        'enable Section field to use the standard security
        'If Not (glbPlantCode = "MISS" Or glbPlantCode = "TROY") Then
        '    clpCode(5).Enabled = False
        'End If
    End If
    
    'Ticket #28174 Franks 05/10/2016
'    chkShowAllEmp.Left = chkShowEmp.Left
'    chkShowAllEmp.Top = chkShowEmp.Top + 300
    chkShowAllEmp.Top = 6920
    chkShowAllEmp.Visible = True
    chkRenewalExce.Top = 6560
    
    'Ticket #28664 Franks 05/30/2016 - WFC not use this funciton, they use chkShowAllEmp
    chkShowEmp.Visible = False
    
    'Ticket #29660 - Contract Employees Enhancement
    If glbWFC Then
        chkExclCONP.Visible = True
        chkExclRET.Visible = True
    Else
        chkExclCONP.Visible = False
        chkExclRET.Visible = False
    End If
End If

'Ticket #21709 Franks 03/12/2012
If glbFormCaption = "Training Plan Report" Then
    Call ScreenSetup4TrainingPlan
End If

'Ticket #29698 - County of Wellington
If glbCompSerial = "S/N - 2262W" Then
    chkReqCourses.Value = 1
    chkLegisCourses.Value = 1
End If

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT

End Sub

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
PosFlag = False: strShift = "": strPosCode = "": strPosGrp = ""

If glbWFC Then
    'Call XLSwriter_WFC
    Call XLSwriter_WFC_New 'Ticket #28174 Franks 05/11/2016
'Commented by Bryan, not done yet.
ElseIf glbCompSerial = "S/N - 2351W" Then 'burlington
    Call XLSwriter_Bur
Else
    'Ticket #22274 - City of Chatham-Kent - Using Training Plan instead of Continuing Education
    If glbCompSerial = "S/N - 2188W" And chkTrainingList Then
        Call XLSwriter_All_TrainList
    Else
        Call XLSwriter_All
    End If
End If

' window title if appropriate
'Me.vbxCrystal.WindowTitle = "Training Matrix Report"

Cri_SetAll = True
Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Training Matrix", "Training Matrix", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Function GetReqCourseCodes()
Dim RsEdEmp As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ, xJobCode
    
    SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
    SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
    SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE " & getWSQLQ(True) & " "
    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    strReqCourses = ","
    Do While Not RsEdEmp.EOF
        SQLQ = "SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 "
        SQLQ = SQLQ & "AND JH_EMPNBR= " & RsEdEmp("ES_EMPNBR")
        xJobCode = ""
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            xJobCode = rsTemp("JH_JOB")
        End If
        rsTemp.Close
        
        If Len(Trim(xJobCode)) > 0 Then
            SQLQ = "SELECT * FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJobCode & "' "
            If rsTemp.State <> 0 Then rsTemp.Close
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsTemp.EOF
            
                strReqCourses = strReqCourses & Trim(rsTemp("PC_CRSCODE")) & ","
                rsTemp.MoveNext
            Loop
            rsTemp.Close
        End If
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    
End Function

Private Sub XLSwriter_WFC_New() 'Ticket #28174 Franks 05/10/2016
'Frank notes:
'the new changes are: if user checked "Show All Employees With or Without Training Records"
'the prorgam will add the employees who don't have any course taken to the list, and then display them on the Excel report
Dim CoJobCodeS As New Collection
Dim CoJobCode As New Collection
Dim rsJobCode As New ADODB.Recordset
Dim RsEdEmp As New ADODB.Recordset
Dim rsCurJob As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strCourseCodes As String
Dim SQLQ, I, J, K, L, M, x, Y, Q, xMax As Integer, xRecNum As Long
Dim exApp As Object, exBook As Object, exSheet As Object
Dim xlsFileTmp, xlsFileMat, xEmpnbr, xNA, xJobDesc, xJobCode
Dim StartLine As Long, strTemp As String
Dim NewDateFormat
Dim CRSDesc(1500) 'Ticket #17410- more than 1000 codes and caused error (1000) '(140) (280)
Dim NAflag(1500) '(1000) '(280)
Dim flgReqC As Boolean, strDisp As String
Dim XLS_Date
Dim rsWRK As New ADODB.Recordset
Dim rsEMP As New ADODB.Recordset
Dim xlocOldNo, xlocCRSCode, xlocOverdueList
Dim xMsg As String
On Error GoTo Err_XLS

    'If glbWFC Then
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainTmpWFC.xls"
    'Else
    '    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainTmp.xls"
    'End If
    'Franks ticket# 6105, put Plant Code in front of "TrainMat.xls", so each plant has own name of Training Matrix report
    xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & clpCode(5) & "TrainMat.xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat
    
    
    'Only the used Jobs under this selection criteria
    If chkCoursesTaken = 0 Then 'If Courses Taken Only checked, don't check which courses are required
                                'Only show the courses which somebody has taken
        Call GetReqCourseCodes
    End If
    'Only the used Jobs under this selection criteria
    
    '''Alway show required courses on the report
    ''SQLQ = "SELECT DISTINCT PC_CRSCODE FROM HR_JOB_COURSE "
    ''rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ''strReqCourses = ","
    ''Do While Not rsTemp.EOF
    ''    strReqCourses = strReqCourses & Trim(rsTemp("PC_CRSCODE")) & ","
    ''    rsTemp.MoveNext
    ''Loop
    ''rsTemp.Close
    '''Alway show required courses on the report
    
    '??? Only show used Course Codes and required codes
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' ORDER BY TB_KEY"
    rsJobCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    I = 1: strCourseCodes = ","
    Do While Not rsJobCode.EOF
        CoJobCodeS.Add Trim(rsJobCode("TB_KEY"))
        CoJobCode.Add I, Trim(rsJobCode("TB_KEY"))
        CRSDesc(I) = Trim(rsJobCode("TB_DESC"))
        'NAflag(i) = 0
        If InStr(1, strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
            NAflag(I) = 1
        Else
            NAflag(I) = 0
        End If
        
        strCourseCodes = strCourseCodes & Trim(rsJobCode("TB_KEY")) & ","
        I = I + 1

        rsJobCode.MoveNext
    Loop
    'I = CoJobCode("1BPR")
    xMax = CoJobCode.count
    rsJobCode.Close
    
  
    'Ticket #20479 Franks 08/29/2011
    'Any course with a renewal date less than today’s date - begin
    If chkRenewalExce.Value Then
        SQLQ = "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        gdbAdoIhr001.Execute SQLQ
        
        SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
        SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
        SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & "WHERE " & getWSQLQ(True) & " "
        SQLQ = SQLQ & "AND NOT (ES_RENEW IS NULL) " ' with renewal date
        SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_CRSCODE,HREDSEM.ES_DATCOMP DESC "
            
        If RsEdEmp.State <> 0 Then RsEdEmp.Close
        RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xRecNum = 0
        If Not RsEdEmp.EOF Then
            K = 0: xRecNum = RsEdEmp.RecordCount
            xlocOldNo = -1
            xlocCRSCode = RsEdEmp("ES_EMPNBR")
            xlocOverdueList = ""
        End If
        Do While Not RsEdEmp.EOF
            MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
            K = K + 1
            DoEvents
            If Not RsEdEmp("ES_CRSCODE") = xlocCRSCode Then
                xlocCRSCode = RsEdEmp("ES_CRSCODE")
                If CVDate(RsEdEmp("ES_RENEW")) < CVDate(Date) Then
                    xlocOverdueList = xlocOverdueList & "'" & Trim(RsEdEmp("ES_CRSCODE")) & "',"
                    'add this employee to the working table
                    Call AddEmpWrokList(RsEdEmp("ES_EMPNBR"))
                End If
            End If
            If Not RsEdEmp("ES_EMPNBR") = xlocOldNo Then
                xlocOldNo = RsEdEmp("ES_EMPNBR")
                xlocCRSCode = "*"
            End If
            RsEdEmp.MoveNext
        Loop
        If Len(xlocOverdueList) > 1 Then
            'remove the last ","
            xlocOverdueList = Left(xlocOverdueList, Len(xlocOverdueList) - 1)
        End If
    End If
    'Any course with a renewal date less than today’s date - end
    
    'To get which columns are all N/A Begin
    SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
    SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
    SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE " & getWSQLQ(True) & " "
    If chkRenewalExce.Value Then 'Ticket #20479 Franks 08/29/2011
        SQLQ = SQLQ & "AND ES_CRSCODE IN (" & xlocOverdueList & ") "
        SQLQ = SQLQ & "AND ES_EMPNBR IN (SELECT TT_EMPNBR FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "') "
    End If
    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    
    If RsEdEmp.State <> 0 Then RsEdEmp.Close
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xRecNum = 0
    xRecNum = RsEdEmp.RecordCount
    K = 0: I = StartLine + 1: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = ""
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        If Len(strPosGrp) > 0 Then 'check Position Group
            If Not IsNull(RsEdEmp("JH_JOB")) Then
                SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & RsEdEmp("JH_JOB") & "' "
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTemp.EOF Then
                    If rsTemp("JB_GRPCD") <> clpCode(14) Then
                        rsTemp.Close
                        GoTo NextLine01
                    End If
                Else
                    rsTemp.Close
                    GoTo NextLine01
                End If
                rsTemp.Close
            End If
        End If
        If Not IsNull(RsEdEmp("ES_CRSCODE")) Then
            If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                'Ticket 4131 Frank 05/27/2003
                If glbWFC Then 'WFC check ES_DATCOMP, if no ES_DATCOMP then don't display it
                    If IsDate(RsEdEmp("ES_DATCOMP")) Then
                        'exSheet.Cells(I, 6 + J) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                        '??? how to get J value
                        'Get J value from CoJobCode
                        J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
                        NAflag(J) = 1
                    End If
                Else
                    NAflag(J) = 1
                End If
            End If
        End If
NextLine01:
        K = K + 1
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    Y = 1
    For x = 1 To xMax
        If NAflag(x) > 0 Then
            NAflag(x) = Y
            Y = Y + 1
        End If
    Next x
    'To get whick columns are all N/A End
    
    'Ticket #21499 Franks 01/26/2012
    'If Y > 280 Then
    If Y > 130 Then
        xMsg = "There are " & Y & " courses taken in this Selection Criteria." & Chr(10)
        xMsg = xMsg & "The report can't show up with more than 130 Course Codes " '280
        'MsgBox "The report can't show up with more than 130 Course Codes" '280
        MsgBox xMsg
        Exit Sub
    End If
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    
    exSheet.Cells(1, 5) = "Training Matrix"
    exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
    exSheet.Cells(2, 1) = "Time: " & Time$
    If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
        exSheet.Cells(2, 5) = "No date entered"
    Else
        strTemp = ""
        If IsDate(dlpDateRange(0).Text) Then
            strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
        End If
        If IsDate(dlpDateRange(1).Text) Then
            strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
        End If
        exSheet.Cells(2, 5) = strTemp
    End If
    
    'Ticket 4131 Frank 05/27/2003
    If glbWFC Then
        StartLine = 6
    Else
        StartLine = 8
    End If
    
    exSheet.Cells(StartLine, 1) = lStr("Division")
    exSheet.Cells(StartLine, 2) = lStr("Department")
    exSheet.Cells(StartLine, 3) = "Employee #"
    exSheet.Cells(StartLine, 4) = "Name"
    exSheet.Cells(StartLine, 5) = "Job Title"
    exSheet.Cells(StartLine, 6) = lStr("Original Hire")
    exSheet.Cells(StartLine, 7) = "Course Taken"
    
    StartLine = StartLine + 1
    Y = 1
    'Display Course Codes and Descaiptions on Title
    For I = 1 To CoJobCode.count
    
        If NAflag(I) > 0 Then
            exSheet.Cells(StartLine, 6 + Y) = CoJobCodeS.Item(I)
            exSheet.Cells(StartLine + 1, 6 + Y) = CRSDesc(I)
            'Ticket 4131 Frank 05/27/2003
            If glbWFC Then
                SQLQ = "SELECT HR_JOB_COURSE.PC_CRSCODE, HR_JOB_COURSE.PC_LEGISLATED FROM HR_JOB_COURSE "
                SQLQ = SQLQ & "WHERE HR_JOB_COURSE.PC_CRSCODE = '" & CoJobCodeS.Item(I) & "' "
                SQLQ = SQLQ & "AND HR_JOB_COURSE.PC_LEGISLATED <> 0 "
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTemp.EOF Then
                    exSheet.Cells(StartLine + 2, 6 + Y) = "Legislated"
                End If
                rsTemp.Close
            Else
                SQLQ = "SELECT HR_JOB_COURSE.PC_CRSCODE, HR_JOB_COURSE.PC_LEGISLATED FROM HR_JOB_COURSE "
                SQLQ = SQLQ & "WHERE HR_JOB_COURSE.PC_CRSCODE = '" & CoJobCodeS.Item(I) & "' "
                'SQLQ = SQLQ & "AND HR_JOB_COURSE.PC_LEGISLATED <> 0 "
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTemp.EOF Then
                    If rsTemp("PC_LEGISLATED") Then
                        exSheet.Cells(StartLine + 2, 6 + Y) = "R/L"
                    Else
                        exSheet.Cells(StartLine + 2, 6 + Y) = "R"
                    End If
                End If
                rsTemp.Close
            End If
            Y = Y + 1
        End If
    Next

    SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
    SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
    SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
    'SQLQ = SQLQ & "WHERE " & getWSQLQ(True) & " "
    'Ticket #21499 Franks 03/08/2012 'fix the problem of "Show All Employees", this not work before
    SQLQ = SQLQ & "WHERE "
    If chkShowEmp Then
        SQLQ = SQLQ & getWSQLQ(False) & " "
    Else
        SQLQ = SQLQ & getWSQLQ(True) & " "
    End If
    
    If chkRenewalExce.Value Then 'Ticket #20479 Franks 08/29/2011
        SQLQ = SQLQ & "AND ES_CRSCODE IN (" & xlocOverdueList & ") "
        SQLQ = SQLQ & "AND ES_EMPNBR IN (SELECT TT_EMPNBR FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "') "
    End If
    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    If RsEdEmp.State <> 0 Then RsEdEmp.Close
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    xRecNum = 0
    If Not RsEdEmp.EOF Then
        RsEdEmp.MoveNext
        RsEdEmp.MoveFirst
        xRecNum = RsEdEmp.RecordCount
    End If
    
    'Ticket #28174 Franks 05/10/2016 - begin
    SQLQ = "DELETE FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.Execute SQLQ
    
    K = 0
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        K = K + 1
        DoEvents
        Call AddEmpToWrk(RsEdEmp, "1")
        RsEdEmp.MoveNext
    Loop
    If Not RsEdEmp.EOF Then
        RsEdEmp.MoveFirst
    End If
    
    If chkShowAllEmp Then
        SQLQ = "SELECT * FROM HREMP WHERE (1=1) AND "
        SQLQ = SQLQ & getWSQLQ(False) & " "
        If rsEMP.State <> 0 Then rsEMP.Close
        rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEMP.EOF Then
            xRecNum = rsEMP.RecordCount
        End If
        K = 0
        Do While Not rsEMP.EOF
            MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
            K = K + 1
            DoEvents
            Call AddEmpToWrk(rsEMP, "2")

            rsEMP.MoveNext
        Loop
    End If
    'Ticket #28174 Franks 05/10/2016 - end
    
    'Open the employee list recordset to show courses - begin
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    ' SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    SQLQ = SQLQ & "ORDER BY TT_NEWDIV,TT_NEWDEPT,TT_NAMEFLD "
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWRK.EOF Then
        xRecNum = rsWRK.RecordCount
    End If
    K = 0: I = StartLine + 2: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = ""
    Y = 1
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        DoEvents
        K = K + 1
        
        'open RsEdEmp --------------- begin
        SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
        SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
        SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & "WHERE HREDSEM.ES_EMPNBR = " & rsWRK("TT_EMPNBR") & " "
        If chkRenewalExce.Value Then 'Ticket #20479 Franks 08/29/2011
            SQLQ = SQLQ & "AND ES_CRSCODE IN (" & xlocOverdueList & ") "
            SQLQ = SQLQ & "AND ES_EMPNBR IN (SELECT TT_EMPNBR FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "') "
        End If
        If IsDate(dlpDateRange(0)) Then SQLQ = SQLQ & " AND ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
        If IsDate(dlpDateRange(1)) Then SQLQ = SQLQ & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
        SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
        If RsEdEmp.State <> 0 Then RsEdEmp.Close
        RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = ""
        
        'If RsEdEmp("ED_EMPNBR") <> xEmpnbr Then ---------- the followings are employee based data
            xEmpnbr = rsWRK("TT_EMPNBR")  'RsEdEmp("ED_EMPNBR")
            
            SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr & " "
            If rsEMP.State <> 0 Then rsEMP.Close
            rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic

            I = I + 1: xNA = 0
            SQLQ = "SELECT JB_CODE, JB_DESCR FROM HRJOB WHERE JB_CODE IN "
            SQLQ = SQLQ & "(SELECT  JH_JOB FROM HR_JOB_HISTORY "
            SQLQ = SQLQ & "WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & xEmpnbr & ") "

            rsCurJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
            xJobDesc = "": xJobCode = ""
            If Not rsCurJob.EOF Then
                If Not IsNull(rsCurJob("JB_DESCR")) Then
                    xJobDesc = rsCurJob("JB_DESCR")
                    xJobCode = rsCurJob("JB_CODE")
                End If
            End If
            rsCurJob.Close
            
            If Not rsEMP.EOF Then
                exSheet.Cells(I, 1) = rsEMP("ED_DIV")
                exSheet.Cells(I, 2) = rsEMP("ED_DEPTNO")
                exSheet.Cells(I, 3) = rsEMP("ED_EMPNBR")
                exSheet.Cells(I, 4) = rsEMP("ED_SURNAME") & "," & rsEMP("ED_FNAME")
                exSheet.Cells(I, 5) = xJobDesc '
                exSheet.Cells(I, 6) = Format(rsEMP("ED_DOH"), NewDateFormat) '"SHORT DATE")
            End If
            rsEMP.Close
            
            'For N/A
            For Q = 1 To xMax
                If NAflag(Q) > 0 Then
                    'exSheet.Cells(i, 6 + Q) = "N/A"
                    exSheet.Cells(I, 6 + NAflag(Q)) = "N/A"
                    Y = Y + 1
                End If
            Next Q
            SQLQ = "SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJobCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsTemp.EOF
                If Not IsNull(rsTemp("PC_CRSCODE")) Then
                    If InStr(1, strCourseCodes, "," & Trim(rsTemp("PC_CRSCODE")) & ",") > 0 Then
                        J = CoJobCode(Trim(rsTemp("PC_CRSCODE")))
                        'Debug.Print NAflag(J), rsTemp("PC_CRSCODE")
                        If NAflag(J) > 0 Then
                            'If the Course Code for this position is Required, then put Blank
                            exSheet.Cells(I, 6 + NAflag(J)) = "" '"N/A"
                            'Ticket 4131 Frank 05/27/2003
                            If Not glbWFC Then
                                exSheet.Cells(I, 6 + NAflag(J) + 124) = "Not Good"
                            End If
                        End If
                        xNA = xNA + 1
                    End If
                End If

                rsTemp.MoveNext
            Loop
            rsTemp.Close

        'End If
        
        If RsEdEmp.EOF Then
            'no couses - do nothing
            'Debug.Print ""
        Else
            Do While Not RsEdEmp.EOF
            'found courses
                If Len(strPosGrp) > 0 Then 'check Position Group
                    If Not IsNull(RsEdEmp("JH_JOB")) Then
                        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & RsEdEmp("JH_JOB") & "' "
                        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                        If Not rsTemp.EOF Then
                            If rsTemp("JB_GRPCD") <> clpCode(14) Then
                                rsTemp.Close
                                GoTo NextLine02
                            End If
                        Else
                            rsTemp.Close
                            GoTo NextLine02
                        End If
                        rsTemp.Close
                    End If
                End If
                If Not IsNull(RsEdEmp("ES_CRSCODE")) Then
                    If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                        J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
                        'Ticket 4131 Frank 05/27/2003
                        If glbWFC Then
                            If IsDate(RsEdEmp("ES_DATCOMP")) Then
                                If NAflag(J) > 0 Then
                                    exSheet.Cells(I, 6 + NAflag(J)) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                                    XLS_Date = RsEdEmp("ES_DATCOMP")
                                    strDisp = ""
                                    If IsDate(RsEdEmp("ES_RENEW")) Then
                                        If DateDiff("d", CVDate(RsEdEmp("ES_DATCOMP")), CVDate(RsEdEmp("ES_RENEW"))) > 365 Then
                                            XLS_Date = RsEdEmp("ES_RENEW")
                                        End If
                                    End If
                                    If DateDiff("d", CVDate(XLS_Date), Now) <= 335 Then
                                        strDisp = "Good"
                                    End If
                                    If DateDiff("d", CVDate(XLS_Date), Now) > 335 And DateDiff("d", CVDate(XLS_Date), Now) <= 365 Then
                                        strDisp = "Expire"
                                    End If
                                    If DateDiff("d", CVDate(XLS_Date), Now) > 335 Then
                                        strDisp = "Not Good"
                                    End If
                                    exSheet.Cells(I, 6 + NAflag(J) + 124) = strDisp
                                    'ticket# 8791 - End
                                End If
                            End If
                        Else 'for Non WFC
                            'Green:   Training in Good Standing
                            'Required Course with no renewal date
                            'Required Course with a renewal date greater than 30 days from today
                            'A non-required course with no renewal date
                            '
                            'Yellow:    Training will Expire in Thirty Days
                            'Any course with a renewal date within 30 days of today
                            '
                            'Red:   Training Not in Good Standing
                            'Required Course with no completed date.
                            'Required Course not taken (No Continuing Education record)
                            'Any course with a renewal date less than today’s date
                        End If
                    End If
    
                End If
                RsEdEmp.MoveNext
            Loop
            
        End If
        'open RsEdEmp --------------- end
       
NextLine02:
        rsWRK.MoveNext
    Loop
    rsWRK.Close
    'Open the employee list recordset to show courses - end
    
    'Exit Sub 'for testing
    
    
    'Save new Excel file as XLS
    'exBook.SaveAs "C:\TrainMat.xls"
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
    Exit Sub
Err_XLS:
'    If Err.Number = 91 Then
'        MsgBox Err.Number
'        Resume Next
'    End If
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriter_WFC_New", "", "Select")
    'Resume Next '???
End Sub

Private Sub XLSwriter_WFC()
Dim CoJobCodeS As New Collection
Dim CoJobCode As New Collection
Dim rsJobCode As New ADODB.Recordset
Dim RsEdEmp As New ADODB.Recordset
Dim rsCurJob As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strCourseCodes As String
Dim SQLQ, I, J, K, L, M, x, Y, Q, xMax As Integer, xRecNum As Long
Dim exApp As Object, exBook As Object, exSheet As Object
Dim xlsFileTmp, xlsFileMat, xEmpnbr, xNA, xJobDesc, xJobCode
Dim StartLine As Long, strTemp As String
Dim NewDateFormat
Dim CRSDesc(1500) 'Ticket #17410- more than 1000 codes and caused error (1000) '(140) (280)
Dim NAflag(1500) '(1000) '(280)
Dim flgReqC As Boolean, strDisp As String
Dim XLS_Date
Dim rsWRK As New ADODB.Recordset
Dim xlocOldNo, xlocCRSCode, xlocOverdueList
Dim xMsg As String
On Error GoTo Err_XLS

    If glbWFC Then
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainTmpWFC.xls"
    Else
        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainTmp.xls"
    End If
    'Franks ticket# 6105, put Plant Code in front of "TrainMat.xls", so each plant has own name of Training Matrix report
    xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & clpCode(5) & "TrainMat.xls"

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat
    
    
    'Only the used Jobs under this selection criteria
    If chkCoursesTaken = 0 Then 'If Courses Taken Only checked, don't check which courses are required
                                'Only show the courses which somebody has taken
        Call GetReqCourseCodes
    End If
    'Only the used Jobs under this selection criteria
    
    '''Alway show required courses on the report
    ''SQLQ = "SELECT DISTINCT PC_CRSCODE FROM HR_JOB_COURSE "
    ''rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ''strReqCourses = ","
    ''Do While Not rsTemp.EOF
    ''    strReqCourses = strReqCourses & Trim(rsTemp("PC_CRSCODE")) & ","
    ''    rsTemp.MoveNext
    ''Loop
    ''rsTemp.Close
    '''Alway show required courses on the report
    
    '??? Only show used Course Codes and required codes
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' ORDER BY TB_KEY"
    rsJobCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    I = 1: strCourseCodes = ","
    Do While Not rsJobCode.EOF
        CoJobCodeS.Add Trim(rsJobCode("TB_KEY"))
        CoJobCode.Add I, Trim(rsJobCode("TB_KEY"))
        CRSDesc(I) = Trim(rsJobCode("TB_DESC"))
        'NAflag(i) = 0
        If InStr(1, strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
            NAflag(I) = 1
        Else
            NAflag(I) = 0
        End If
        
        strCourseCodes = strCourseCodes & Trim(rsJobCode("TB_KEY")) & ","
        I = I + 1

        rsJobCode.MoveNext
    Loop
    'I = CoJobCode("1BPR")
    xMax = CoJobCode.count
    rsJobCode.Close
    
    'Ticket #20479 Franks 08/29/2011
    'Any course with a renewal date less than today’s date - begin
    If chkRenewalExce.Value Then
        SQLQ = "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "' "
        gdbAdoIhr001.Execute SQLQ
        
        SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
        SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
        SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & "WHERE " & getWSQLQ(True) & " "
        SQLQ = SQLQ & "AND NOT (ES_RENEW IS NULL) " ' with renewal date
        SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_CRSCODE,HREDSEM.ES_DATCOMP DESC "
            
        If RsEdEmp.State <> 0 Then RsEdEmp.Close
        RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xRecNum = 0
        If Not RsEdEmp.EOF Then
            K = 0: xRecNum = RsEdEmp.RecordCount
            xlocOldNo = -1
            xlocCRSCode = RsEdEmp("ES_EMPNBR")
            xlocOverdueList = ""
        End If
        Do While Not RsEdEmp.EOF
            MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
            K = K + 1
            DoEvents
            If Not RsEdEmp("ES_CRSCODE") = xlocCRSCode Then
                xlocCRSCode = RsEdEmp("ES_CRSCODE")
                If CVDate(RsEdEmp("ES_RENEW")) < CVDate(Date) Then
                    xlocOverdueList = xlocOverdueList & "'" & Trim(RsEdEmp("ES_CRSCODE")) & "',"
                    'add this employee to the working table
                    Call AddEmpWrokList(RsEdEmp("ES_EMPNBR"))
                End If
            End If
            If Not RsEdEmp("ES_EMPNBR") = xlocOldNo Then
                xlocOldNo = RsEdEmp("ES_EMPNBR")
                xlocCRSCode = "*"
            End If
            RsEdEmp.MoveNext
        Loop
        If Len(xlocOverdueList) > 1 Then
            'remove the last ","
            xlocOverdueList = Left(xlocOverdueList, Len(xlocOverdueList) - 1)
        End If
    End If
    'Any course with a renewal date less than today’s date - end
    
    'To get which columns are all N/A Begin
    SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
    SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
    SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
    SQLQ = SQLQ & "WHERE " & getWSQLQ(True) & " "
    If chkRenewalExce.Value Then 'Ticket #20479 Franks 08/29/2011
        SQLQ = SQLQ & "AND ES_CRSCODE IN (" & xlocOverdueList & ") "
        SQLQ = SQLQ & "AND ES_EMPNBR IN (SELECT TT_EMPNBR FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "') "
    End If
    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    
    If RsEdEmp.State <> 0 Then RsEdEmp.Close
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xRecNum = 0
    xRecNum = RsEdEmp.RecordCount
    K = 0: I = StartLine + 1: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = ""
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        If Len(strPosGrp) > 0 Then 'check Position Group
            If Not IsNull(RsEdEmp("JH_JOB")) Then
                SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & RsEdEmp("JH_JOB") & "' "
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTemp.EOF Then
                    If rsTemp("JB_GRPCD") <> clpCode(14) Then
                        rsTemp.Close
                        GoTo NextLine01
                    End If
                Else
                    rsTemp.Close
                    GoTo NextLine01
                End If
                rsTemp.Close
            End If
        End If
        If Not IsNull(RsEdEmp("ES_CRSCODE")) Then
            If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                'Ticket 4131 Frank 05/27/2003
                If glbWFC Then 'WFC check ES_DATCOMP, if no ES_DATCOMP then don't display it
                    If IsDate(RsEdEmp("ES_DATCOMP")) Then
                        'exSheet.Cells(I, 6 + J) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                        '??? how to get J value
                        'Get J value from CoJobCode
                        J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
                        NAflag(J) = 1
                    End If
                Else
                    NAflag(J) = 1
                End If
            End If
        End If
NextLine01:
        K = K + 1
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    Y = 1
    For x = 1 To xMax
        If NAflag(x) > 0 Then
            NAflag(x) = Y
            Y = Y + 1
        End If
    Next x
    'To get whick columns are all N/A End
    
    'Ticket #21499 Franks 01/26/2012
    'If Y > 280 Then
    If Y > 130 Then
        xMsg = "There are " & Y & " courses taken in this Selection Criteria." & Chr(10)
        xMsg = xMsg & "The report can't show up with more than 130 Course Codes " '280
        'MsgBox "The report can't show up with more than 130 Course Codes" '280
        MsgBox xMsg
        Exit Sub
    End If
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    
    exSheet.Cells(1, 5) = "Training Matrix"
    exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
    exSheet.Cells(2, 1) = "Time: " & Time$
    If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
        exSheet.Cells(2, 5) = "No date entered"
    Else
        strTemp = ""
        If IsDate(dlpDateRange(0).Text) Then
            strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
        End If
        If IsDate(dlpDateRange(1).Text) Then
            strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
        End If
        exSheet.Cells(2, 5) = strTemp
    End If
    
    'Ticket 4131 Frank 05/27/2003
    If glbWFC Then
        StartLine = 6
    Else
        StartLine = 8
    End If
    
    exSheet.Cells(StartLine, 1) = lStr("Division")
    exSheet.Cells(StartLine, 2) = lStr("Department")
    exSheet.Cells(StartLine, 3) = "Employee #"
    exSheet.Cells(StartLine, 4) = "Name"
    exSheet.Cells(StartLine, 5) = "Job Title"
    exSheet.Cells(StartLine, 6) = lStr("Original Hire")
    exSheet.Cells(StartLine, 7) = "Course Taken"
    
    StartLine = StartLine + 1
    Y = 1
    'Display Course Codes and Descaiptions on Title
    For I = 1 To CoJobCode.count
    
        If NAflag(I) > 0 Then
            exSheet.Cells(StartLine, 6 + Y) = CoJobCodeS.Item(I)
            exSheet.Cells(StartLine + 1, 6 + Y) = CRSDesc(I)
            'Ticket 4131 Frank 05/27/2003
            If glbWFC Then
                SQLQ = "SELECT HR_JOB_COURSE.PC_CRSCODE, HR_JOB_COURSE.PC_LEGISLATED FROM HR_JOB_COURSE "
                SQLQ = SQLQ & "WHERE HR_JOB_COURSE.PC_CRSCODE = '" & CoJobCodeS.Item(I) & "' "
                SQLQ = SQLQ & "AND HR_JOB_COURSE.PC_LEGISLATED <> 0 "
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTemp.EOF Then
                    exSheet.Cells(StartLine + 2, 6 + Y) = "Legislated"
                End If
                rsTemp.Close
            Else
                SQLQ = "SELECT HR_JOB_COURSE.PC_CRSCODE, HR_JOB_COURSE.PC_LEGISLATED FROM HR_JOB_COURSE "
                SQLQ = SQLQ & "WHERE HR_JOB_COURSE.PC_CRSCODE = '" & CoJobCodeS.Item(I) & "' "
                'SQLQ = SQLQ & "AND HR_JOB_COURSE.PC_LEGISLATED <> 0 "
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTemp.EOF Then
                    If rsTemp("PC_LEGISLATED") Then
                        exSheet.Cells(StartLine + 2, 6 + Y) = "R/L"
                    Else
                        exSheet.Cells(StartLine + 2, 6 + Y) = "R"
                    End If
                End If
                rsTemp.Close
            End If
            Y = Y + 1
        End If
    Next

    SQLQ = "SELECT HREDSEM.*, HREMP.ED_EMPNBR,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DIV, "
    SQLQ = SQLQ & "HREMP.ED_DEPTNO,HREMP.ED_DOH "
    SQLQ = SQLQ & "FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
    'SQLQ = SQLQ & "WHERE " & getWSQLQ(True) & " "
    'Ticket #21499 Franks 03/08/2012 'fix the problem of "Show All Employees", this not work before
    SQLQ = SQLQ & "WHERE "
    If chkShowEmp Then
        SQLQ = SQLQ & getWSQLQ(False) & " "
    Else
        SQLQ = SQLQ & getWSQLQ(True) & " "
    End If
    
    If chkRenewalExce.Value Then 'Ticket #20479 Franks 08/29/2011
        SQLQ = SQLQ & "AND ES_CRSCODE IN (" & xlocOverdueList & ") "
        SQLQ = SQLQ & "AND ES_EMPNBR IN (SELECT TT_EMPNBR FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "') "
    End If
    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    If RsEdEmp.State <> 0 Then RsEdEmp.Close
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    xRecNum = 0
    If Not RsEdEmp.EOF Then
        RsEdEmp.MoveNext
        RsEdEmp.MoveFirst
        xRecNum = RsEdEmp.RecordCount
    End If
    
    K = 0: I = StartLine + 2: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = ""
    Y = 1
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        If Len(strPosGrp) > 0 Then 'check Position Group
            If Not IsNull(RsEdEmp("JH_JOB")) Then
                SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & RsEdEmp("JH_JOB") & "' "
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsTemp.EOF Then
                    If rsTemp("JB_GRPCD") <> clpCode(14) Then
                        rsTemp.Close
                        GoTo NextLine02
                    End If
                Else
                    rsTemp.Close
                    GoTo NextLine02
                End If
                rsTemp.Close
            End If
        End If
        If RsEdEmp("ED_EMPNBR") <> xEmpnbr Then
            xEmpnbr = RsEdEmp("ED_EMPNBR")

            I = I + 1: xNA = 0
            SQLQ = "SELECT JB_CODE, JB_DESCR FROM HRJOB WHERE JB_CODE IN "
            SQLQ = SQLQ & "(SELECT  JH_JOB FROM HR_JOB_HISTORY "
            SQLQ = SQLQ & "WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & xEmpnbr & ") "
            
            rsCurJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
            xJobDesc = "": xJobCode = ""
            If Not rsCurJob.EOF Then
                If Not IsNull(rsCurJob("JB_DESCR")) Then
                    xJobDesc = rsCurJob("JB_DESCR")
                    xJobCode = rsCurJob("JB_CODE")
                End If
            End If
            rsCurJob.Close
            
            exSheet.Cells(I, 1) = RsEdEmp("ED_DIV")
            exSheet.Cells(I, 2) = RsEdEmp("ED_DEPTNO")
            exSheet.Cells(I, 3) = RsEdEmp("ED_EMPNBR")
            exSheet.Cells(I, 4) = RsEdEmp("ED_SURNAME") & "," & RsEdEmp("ED_FNAME")
            exSheet.Cells(I, 5) = xJobDesc '
            exSheet.Cells(I, 6) = Format(RsEdEmp("ED_DOH"), NewDateFormat) '"SHORT DATE")

            'For N/A
            For Q = 1 To xMax
                If NAflag(Q) > 0 Then
                    'exSheet.Cells(i, 6 + Q) = "N/A"
                    exSheet.Cells(I, 6 + NAflag(Q)) = "N/A"
                    Y = Y + 1
                End If
            Next Q
            SQLQ = "SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJobCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsTemp.EOF
                If Not IsNull(rsTemp("PC_CRSCODE")) Then
                    If InStr(1, strCourseCodes, "," & Trim(rsTemp("PC_CRSCODE")) & ",") > 0 Then
                        J = CoJobCode(Trim(rsTemp("PC_CRSCODE")))
                        'Debug.Print NAflag(J), rsTemp("PC_CRSCODE")
                        If NAflag(J) > 0 Then
                            'If the Course Code for this position is Required, then put Blank
                            exSheet.Cells(I, 6 + NAflag(J)) = "" '"N/A"
                            'Ticket 4131 Frank 05/27/2003
                            If Not glbWFC Then
                                exSheet.Cells(I, 6 + NAflag(J) + 124) = "Not Good"
                            End If
                        End If
                        xNA = xNA + 1
                    End If
                End If

                rsTemp.MoveNext
            Loop
            rsTemp.Close
            
        End If
        
        If Not IsNull(RsEdEmp("ES_CRSCODE")) Then
            If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
                'Ticket 4131 Frank 05/27/2003
                If glbWFC Then
                    If IsDate(RsEdEmp("ES_DATCOMP")) Then
                        ''exSheet.Cells(i, 6 + J) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                        
                        'ticket# 8791 - Begin
                        'If IsDate(RsEdEmp("ES_RENEW")) Then 'Ticket 6587 item 18
                        '    exSheet.Cells(I, 6 + NAflag(J)) = Format(RsEdEmp("ES_RENEW"), NewDateFormat)
                        'Else
                        '    exSheet.Cells(I, 6 + NAflag(J)) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                        'End If
                        exSheet.Cells(I, 6 + NAflag(J)) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                        XLS_Date = RsEdEmp("ES_DATCOMP")
                        strDisp = ""
                        If IsDate(RsEdEmp("ES_RENEW")) Then
                            If DateDiff("d", CVDate(RsEdEmp("ES_DATCOMP")), CVDate(RsEdEmp("ES_RENEW"))) > 365 Then
                                XLS_Date = RsEdEmp("ES_RENEW")
                            End If
                        End If
                        If DateDiff("d", CVDate(XLS_Date), Now) <= 335 Then
                            strDisp = "Good"
                        End If
                        If DateDiff("d", CVDate(XLS_Date), Now) > 335 And DateDiff("d", CVDate(XLS_Date), Now) <= 365 Then
                            strDisp = "Expire"
                        End If
                        If DateDiff("d", CVDate(XLS_Date), Now) > 335 Then
                            strDisp = "Not Good"
                        End If
                        exSheet.Cells(I, 6 + NAflag(J) + 124) = strDisp
                        'ticket# 8791 - End
                        
                    End If
                Else
                    'Green:   Training in Good Standing
                    'Required Course with no renewal date
                    'Required Course with a renewal date greater than 30 days from today
                    'A non-required course with no renewal date
                    '
                    'Yellow:    Training will Expire in Thirty Days
                    'Any course with a renewal date within 30 days of today
                    '
                    'Red:   Training Not in Good Standing
                    'Required Course with no completed date.
                    'Required Course not taken (No Continuing Education record)
                    'Any course with a renewal date less than today’s date
                    
                    'Check If Required Courses
                    If IsDate(RsEdEmp("ES_DATCOMP")) Then
                        exSheet.Cells(I, 6 + NAflag(J)) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                    Else
                        exSheet.Cells(I, 6 + NAflag(J)) = ""
                    End If
                    strDisp = "Good"
                    If InStr(1, strReqCourses, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                        flgReqC = True
                    Else
                        flgReqC = False
                    End If
                    'for Good Standing - Green
                    If flgReqC And Not IsDate(RsEdEmp("ES_RENEW")) Then
                        strDisp = "Good"
                    End If
                    If flgReqC And IsDate(RsEdEmp("ES_RENEW")) Then
                        If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) > 30 Then
                            strDisp = "Good"
                        End If
                    End If
                    If Not flgReqC And Not IsDate(RsEdEmp("ES_RENEW")) Then
                        strDisp = "Good"
                    End If
                    'Yellow:    Training will Expire in Thirty Days
                    If IsDate(RsEdEmp("ES_RENEW")) Then
                        If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) >= 0 And DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) <= 30 Then
                            strDisp = "Expire"
                        End If
                    End If
                    'Red:   Training Not in Good Standing
                    If IsDate(RsEdEmp("ES_RENEW")) Then
                        If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) < 0 Then
                            strDisp = "Not Good"
                        End If
                    End If
                    If flgReqC And Not IsDate(RsEdEmp("ES_DATCOMP")) Then
                        strDisp = "Not Good"
                    End If
                    exSheet.Cells(I, 6 + NAflag(J) + 124) = strDisp
                End If
            End If
        End If
NextLine02:
        K = K + 1
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    
    
    'Save new Excel file as XLS
    'exBook.SaveAs "C:\TrainMat.xls"
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
    Exit Sub
Err_XLS:
'    If Err.Number = 91 Then
'        MsgBox Err.Number
'        Resume Next
'    End If
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If

    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriter_WFC", "", "Select")

End Sub

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function
'
'Private Sub Cri_Div()
'
'Dim DivCri As String
'Dim countr   As Integer  ' EEList_Snap is definded at form level
'
'
'If Len(clpDiv.Text) > 0 Then
'    DivCri = "((HREMP.ED_DIV) in ['" & REplace(clpDiv.Text, ",", "','") & "'])"
'End If
'
'If Len(DivCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = DivCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & DivCri
'    End If
'    glbiOneWhere = True
'End If
'
'End Sub

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

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%, fuTempCri As String

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_TRAIN.TR_RENEW} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    fuTempCri = "(TR_RENEW >=" & Date_SQL(dlpDateRange(0).Text)
    fuTempCri = fuTempCri & " AND TR_RENEW <= " & Date_SQL(dlpDateRange(1).Text) & ")"
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(x%).Text) > 0 Then
        TempCri = "({HR_TRAIN.TR_RENEW}  "
        If x% = 0 Then
            TempCri = TempCri & " >= "
            fuTempCri = "(TR_RENEW >= " & Date_SQL(dlpDateRange(0).Text) & ") "
        Else
            TempCri = TempCri & " <= "
            fuTempCri = "(TR_RENEW <= " & Date_SQL(dlpDateRange(1).Text) & ") "
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
    'DateSelCri = IIf(Len(DateSelCri) > 0, DateSelCri & " AND ", "") & fuTempCri
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "({HREMP.ED_PT} IN ['" & Replace(clpPT.Text, ",", "','") & "'])"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

'Private Sub Cri_Shift()
'Dim EECri As String, OneSet%, X%
'
'    'Looking for Shift from Job_History
'    'If Len(txtShift.Text) < 1 Then Exit Sub
'    'EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"
'    '
'    'If glbiOneWhere Then
'    '    glbstrSelCri = glbstrSelCri & " AND " & EECri
'    'Else
'    '    glbstrSelCri = EECri
'    'End If
'    'glbiOneWhere = True
'
'    If Len(txtShift.Text) < 1 Then Exit Sub
'    strShift = "((HR_JOB_HISTORY.JH_SHIFT)='" & txtShift & "')"
'    PosFlag = True
'End Sub
'Private Sub Cri_PosCode()
'    If Len(clpJob.Text) < 1 Then Exit Sub
'    strPosCode = "((HR_JOB_HISTORY.JH_JOB)='" & clpJob.Text & "')"
'    PosFlag = True
'End Sub
'Private Sub Cri_PosGrp()
'    If Len(clpCode(14).Text) < 1 Then Exit Sub
'    strPosGrp = "((HRJOB.JB_GRPCD)='" & clpCode(14).Text & "')"
'    PosFlag = True
'End Sub
'Private Sub Cri_Code(intIdx%)
'Dim CodeCri As String
'Dim countr   As Integer  ' EEList_Snap is definded at form level
'Dim strCd$
'
'If Len(clpCode(intIdx%).Text) > 0 Then
'    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
'    If intIdx% = 1 Then strCd$ = "HREMP.ED_ORG"
'    If intIdx% = 2 Then strCd$ = "HREMP.ED_EMP"
'    If intIdx% = 3 Then strCd$ = "HREMP.ED_REGION"
'    If intIdx% = 4 Then strCd$ = "HREMP.ED_ADMINBY"
'    If intIdx% = 5 Then strCd$ = "HREMP.ED_SECTION"
'    CodeCri = "((" & strCd$ & ") = '" & clpCode(intIdx%).Text & "')"
'    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
'        CodeCri = "(((" & strCd$ & ") = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ((" & strCd$ & ") = 'ALL" & clpCode(intIdx%).Text & "') )"
'    End If
'End If
'
'If Len(CodeCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = CodeCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & CodeCri
'    End If
'    glbiOneWhere = True
'End If
''
''If Len(clpCode(0)) <> 0 Then SQLQ = SQLQ & " AND ED_LOC='" & clpCode(0) & "'"
''If Len(clpCode(1)) <> 0 Then SQLQ = SQLQ & " AND ED_ORG='" & clpCode(0) & "'"
''If Len(clpCode(2)) <> 0 Then SQLQ = SQLQ & " AND ED_EMP='" & clpCode(0) & "'"
''If Len(clpCode(3)) <> 0 Then SQLQ = SQLQ & " AND ED_REGION='" & clpCode(0) & "'"
''If Len(clpCode(4)) <> 0 Then SQLQ = SQLQ & " AND ED_ADMINBY='" & clpCode(0) & "'"
''If Len(clpCode(5)) <> 0 Then SQLQ = SQLQ & " AND ED_SECTION='" & clpCode(0) & "'"
''
'
'
'End Sub

Private Function CriCheck()
Dim x%

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

For x% = 0 To 5
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

If Len(clpJob.Text) > 0 And clpJob.Caption = "Unassigned" Then
    MsgBox "Job code must be valid"
    clpJob.SetFocus
    Exit Function
End If

'Hemu - 05/13/2003 Begin - From Date and To Date
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If
'Hemu - 05/13/2003 End

'If glbWFC Then
If glbWFC And glbFormCaption = "Training Matrix Report" Then
    If Len(clpCode(5)) = 0 Then
        MsgBox lStr("Section is required.")
        clpCode(5).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDateRange(0)) Then
        MsgBox "From Date is required!"                       '
        Me.dlpDateRange(0).SetFocus
        Exit Function
    End If
    If Not IsDate(dlpDateRange(1)) Then
        MsgBox "To Date is required!"                       '
        Me.dlpDateRange(1).SetFocus
        Exit Function
    End If
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If clpProv.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Residence"
    clpProv.SetFocus
    Exit Function
End If

If clpProvEmp.Caption = "Unassigned" Then
    MsgBox "Invalid Prov. of Employment"
    clpProvEmp.SetFocus
    Exit Function
End If

CriCheck = True

End Function

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Private Sub txtShift_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub Cri_CourseCode()
Dim EECri As String, OneSet%, x%
Dim strC2, strCx As String
Dim strCa$


If Len(clpCrsCode.Text) < 1 Then Exit Sub

strCa$ = "HR_TRAIN.TR_CRSCODE"
EECri = "({" & strCa$ & "} in ['" & Replace(clpCrsCode.Text, ",", "','") & "'])"
 
If Len(EECri) > 0 Then
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & "(" & EECri & ")"
    Else
        glbstrSelCri = EECri
    End If
    glbiOneWhere = True
End If
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

'Hemu - 08/25/2003 Begin
Private Sub XLSwriter_All()
Dim CoJobCodeS As New Collection
Dim CoJobCode As New Collection
Dim rsJobCode As New ADODB.Recordset
Dim RsEdEmp As New ADODB.Recordset
Dim rsCurJob As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strCourseCodes As String
Dim SQLQ, I, J, K, L, M, x, Y, Q, xMax As Integer, xRecNum As Long
Dim exApp As Object, exBook As Object, exSheet As Object
Dim xlsFileTmp, xlsFileMat, xEmpnbr, xNA, xJobDesc, xJobCode, xLocation
Dim StartLine As Long, strTemp As String
Dim NewDateFormat
Dim CRSDesc(260)
Dim NAflag(260)
Dim flgReqC As Boolean, strDisp As String  'strReqCourses As String,

Dim z As Integer
Dim strCourse As String
Dim strAndWhere As String
Dim QStr, QStr1
Dim xRes

On Error GoTo Err_XLS

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainTmp.xls"
    'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainMat.xls"
    'Ticket# 8293
    If glbLinamar Then 'Or glbCompSerial = "S/N - 2336W" Then
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Reports\TrainMat" & Trim(glbUserID) & ".xls"
    Else
        xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "TrainMat" & Trim(glbUserID) & ".xls"
    End If

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat
    'Get Required Course....
    
    SQLQ = "SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE 1=1 "
    SQLQ = SQLQ & " AND PC_JOB IN (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & getWSQLQ(False) & "))"
    If clpJob <> "" Then SQLQ = SQLQ & " AND PC_JOB = '" & clpJob.Text & "' "
    If clpCrsCode <> "" Then SQLQ = SQLQ & " AND PC_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
    If chkReqCourses Then If chkLegisCourses Then SQLQ = SQLQ & " AND PC_LEGISLATED <> 0"

    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    strReqCourses = ","
    Do While Not rsTemp.EOF
        strReqCourses = strReqCourses & Trim(rsTemp("PC_CRSCODE")) & ","
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    'Get Course From HRTABL
    'Hemu - Since if there are more than 260 courses in the table it gives the Subscript out of range error
    '       and also Excel cannot take more than 124cols, the user may select only the courses
    '       needed in the report, made the following change to only select from the HRTABL the
    '       courses based on the selection criteria given by the user
    
    QStr = ""
    QStr1 = ""
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' "
    'If clpCrsCode <> "" Then SQLQ = SQLQ & " AND TB_KEY IN ('" & Replace(clpCrsCode, ",", "','") & "') "
If (chkReqCourses.Value = 0) Then
    SQLQ = SQLQ & " AND TB_KEY IN (SELECT ES_CRSCODE FROM HREDSEM WHERE ES_EMPNBR IN"
    SQLQ = SQLQ & " (SELECT ED_EMPNBR FROM HREMP "
    
    Call glbCri_DeptUN(clpDept.Text)

    QStr = Replace(Replace(glbstrSelCri, "{", "("), "}", ")")
    If glbSQL Or glbOracle Then
        QStr = Replace(Replace(QStr, "[", "("), "]", ")")
    End If
    If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
    If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"
    If clpCode(1) <> "" Then
        'QStr = QStr & " AND ED_ORG='" & clpCode(1) & "'"
        QStr = QStr & " AND ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
    End If
    If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
    If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
    If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
    If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
    If clpPT <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpPT, ",", "','") & "')"
    If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
        
    If clpProv <> "" Then QStr = QStr & " AND ED_PROV='" & clpProv & "'"
    If clpProvEmp <> "" Then QStr = QStr & " AND ED_PROVEMP='" & clpProv & "'"
        
    'Ticket #29660 - Contract Employees Enhancement
    If glbWFC Then
        If chkExclCONP.Visible And chkExclRET.Visible = True Then
            If chkExclCONP Then
                QStr = QStr & " AND ED_EMP <> 'CONP'"
            End If
            If chkExclRET Then
                QStr = QStr & " AND ED_EMP <> 'RET'"
            End If
        End If
    End If
        
    If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
        QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
        If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
        If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
        If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
        QStr = QStr & ")"
    End If
                
    If IsDate(dlpDateRange(0)) Then QStr1 = QStr1 & " AND ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
    If IsDate(dlpDateRange(1)) Then QStr1 = QStr1 & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
    If clpCrsType <> "" Then QStr1 = QStr1 & " AND ES_CTYPE IN ('" & Replace(clpCrsType, ",", "','") & "') "
    If clpCrsCode <> "" Then QStr1 = QStr1 & " AND ES_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
    QStr = QStr & ")" & QStr1
    
    SQLQ = SQLQ & " WHERE " & QStr
    SQLQ = SQLQ & ")"
End If
SQLQ = SQLQ & " ORDER BY TB_KEY"
    
    rsJobCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    I = 1: strCourseCodes = ","
    Do While Not rsJobCode.EOF
    
        'If (glbCompSerial = "S/N - 2288W" And chkReqCourses) Then
        If (chkReqCourses) Then
            If InStr(strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
                CoJobCodeS.Add Trim(rsJobCode("TB_KEY"))
                CoJobCode.Add I, Trim(rsJobCode("TB_KEY"))
                CRSDesc(I) = Trim(rsJobCode("TB_DESC"))
            End If
        Else
            CoJobCodeS.Add Trim(rsJobCode("TB_KEY"))
            CoJobCode.Add I, Trim(rsJobCode("TB_KEY"))
            CRSDesc(I) = Trim(rsJobCode("TB_DESC"))
        End If
        
        If InStr(strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
            NAflag(I) = 1
        Else
            NAflag(I) = 0
        End If

        If (chkReqCourses) Then
            If InStr(1, strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
                strCourseCodes = strCourseCodes & Trim(rsJobCode("TB_KEY")) & ","
                I = I + 1
            End If
        Else
            strCourseCodes = strCourseCodes & Trim(rsJobCode("TB_KEY")) & ","
            I = I + 1
        End If
        rsJobCode.MoveNext
    Loop
    xMax = CoJobCode.count
    rsJobCode.Close
    
    If xMax > 124 Then
        xRes = MsgBox("Course Code exceeds 124 courses: Training Matrix report will not fit in MS Excel spreadsheet." & Chr(10) & "Use the Selection Criteria to narrow down the courses." & Chr(10) & Chr(10) & "Exiting Training Matrix report.", vbOKOnly, "info:HR - Course Code exceeds 124 columns")
        Exit Sub
    End If
    
    'To get which columns are all N/A Begin
'    If Not PosFlag Then 'Len(txtShift) = 0 Then
    If Not glbOracle Then
        SQLQ = "SELECT ES_CRSCODE "
        SQLQ = SQLQ & " FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE ES_CRSCODE IS NOT NULL "
        SQLQ = SQLQ & " AND " & getWSQLQ(True) & " "
    Else
        SQLQ = "SELECT ES_CRSCODE "
        SQLQ = SQLQ & " FROM HREDSEM,HREMP "
        SQLQ = SQLQ & " WHERE  ES_CRSCODE IS NOT NULL "
        SQLQ = SQLQ & " AND " & getWSQLQ(True) & " "
        SQLQ = SQLQ & " AND (HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR)"
    End If
    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xRecNum = 0
    xRecNum = RsEdEmp.RecordCount
    K = 0: I = StartLine + 1: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = "": xLocation = ""
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
           J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
           NAflag(J) = 1
        End If
NextLine01:
        K = K + 1
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    Y = 1
    For x = 1 To xMax
        If NAflag(x) > 0 Then 'How many course codes display and on which field in xls
            NAflag(x) = Y
            Y = Y + 1
        End If
    Next x
    'To get whick columns are all N/A End
    
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    
    exSheet.Cells(1, 5) = "Training Matrix"
    exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
    exSheet.Cells(2, 1) = "Time: " & Time$
    If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
        exSheet.Cells(2, 5) = "No date entered"
    Else
        strTemp = ""
        If IsDate(dlpDateRange(0).Text) Then
            strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
        End If
        If IsDate(dlpDateRange(1).Text) Then
            strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
        End If
        exSheet.Cells(2, 5) = strTemp
    End If
    'Ticket 4131 Frank 05/27/2003
    If glbWFC Then
        StartLine = 6
    Else
        StartLine = 8
    End If
    
    exSheet.Cells(StartLine, 1) = lStr("Division")
    exSheet.Cells(StartLine, 2) = lStr("Department")
    exSheet.Cells(StartLine, 3) = "Employee #"
    exSheet.Cells(StartLine, 4) = "Name"
    exSheet.Cells(StartLine, 5) = "Job Title"
    'Zahoor(Sam) 03/02/2006
    exSheet.Cells(StartLine, 6) = "Location"
    'Zahoor(Sam) 03/02/2006
    exSheet.Cells(StartLine, 7) = lStr("Original Hire")
    exSheet.Cells(StartLine, 8) = "Course Taken"
    
    
    StartLine = StartLine + 1
    Y = 1
    'Display Course Codes and Descriptions on Title
    For I = 1 To CoJobCode.count
        If NAflag(I) > 0 Then
            exSheet.Cells(StartLine, 7 + Y) = CoJobCodeS.Item(I)
            exSheet.Cells(StartLine + 1, 7 + Y) = CRSDesc(I)

            SQLQ = "SELECT HR_JOB_COURSE.PC_CRSCODE, HR_JOB_COURSE.PC_JOB, HR_JOB_COURSE.PC_LEGISLATED FROM HR_JOB_COURSE "
            SQLQ = SQLQ & "WHERE HR_JOB_COURSE.PC_CRSCODE = '" & CoJobCodeS.Item(I) & "' "
            SQLQ = SQLQ & " AND PC_JOB IN (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & getWSQLQ(False) & "))"
            If clpJob <> "" Then SQLQ = SQLQ & " AND PC_JOB = '" & clpJob.Text & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                If rsTemp("PC_LEGISLATED") Then
                    exSheet.Cells(StartLine + 2, 7 + Y) = "R/L"
                Else
                    exSheet.Cells(StartLine + 2, 7 + Y) = "R"
                End If
            End If
            rsTemp.Close
            Y = Y + 1
        End If
    Next
    SQLQ = "SELECT ES_CRSCODE, ES_DATCOMP, ES_RENEW, "
    'Zahoor(Sam) 03/02/2006
    'old query
    'SQLQ = SQLQ & " ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DIV, ED_DEPTNO,ED_DOH "
    
    'new query added ED_LOC on the request of City of Chatham-Kent -tkt #2188 request
    SQLQ = SQLQ & " ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DIV, ED_DEPTNO,ED_DOH,ED_LOC "
    'Zahoor(Sam) 03/02/2006
    
    If Not glbOracle Then
        SQLQ = SQLQ & " FROM HREMP LEFT JOIN HREDSEM ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE "
    Else
        SQLQ = SQLQ & " FROM HREDSEM,HREMP "
        'Ticket #15688 - Begin 'wrong direction of left join for Oracle
        'SQLQ = SQLQ & " WHERE (HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR(+))"
        SQLQ = SQLQ & " WHERE (HREMP.ED_EMPNBR = HREDSEM.ES_EMPNBR(+))"
        'Ticket #15688 - End
        SQLQ = SQLQ & " AND "
    End If
    If chkShowEmp Then
        SQLQ = SQLQ & getWSQLQ(False) & " "
    Else
        SQLQ = SQLQ & getWSQLQ(True) & " "
    End If
    SQLQ = SQLQ & "ORDER BY ED_DIV,ED_DEPTNO,ED_SURNAME,ED_FNAME,ES_DATCOMP "
    
    If RsEdEmp.State <> 0 Then RsEdEmp.Close
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    xRecNum = RsEdEmp.RecordCount
    
    K = 0: I = StartLine + 2: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = "": xLocation = ""
    Y = 1
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        If RsEdEmp("ED_EMPNBR") <> xEmpnbr Then
            xEmpnbr = RsEdEmp("ED_EMPNBR")

            I = I + 1: xNA = 0
            
            SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB "
            SQLQ = SQLQ & " WHERE JB_CODE IN (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & xEmpnbr & ")"
            rsCurJob.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            xJobDesc = "": xJobCode = ""
            If Not rsCurJob.EOF Then
                If Not IsNull(rsCurJob("JB_DESCR")) Then
                    xJobDesc = rsCurJob("JB_DESCR")
                    xJobCode = rsCurJob("JB_CODE")
                End If
            End If
            rsCurJob.Close
            
            exSheet.Cells(I, 1) = RsEdEmp("ED_DIV")
            exSheet.Cells(I, 2) = RsEdEmp("ED_DEPTNO")
            exSheet.Cells(I, 3) = RsEdEmp("ED_EMPNBR")
            exSheet.Cells(I, 4) = RsEdEmp("ED_SURNAME") & "," & RsEdEmp("ED_FNAME")
            exSheet.Cells(I, 5) = xJobDesc '
           ''Zahoor(Sam) 03/02/2006
            exSheet.Cells(I, 6) = RsEdEmp("ED_LOC")
            'Zahoor(Sam) 03/02/2006
            'Ticket #24875 - Day and Month switching in Excel if Day <= 12
            If UCase(NewDateFormat) <> "MM/DD/YYYY" And UCase(NewDateFormat) <> "DD/MM/YYYY" Then
                exSheet.Cells(I, 7) = Format(RsEdEmp("ED_DOH"), NewDateFormat) '"SHORT DATE")
            Else
                exSheet.Cells(I, 7) = Format(RsEdEmp("ED_DOH"), "mm/dd/yyyy") '"SHORT DATE")
            End If
            
           
            
            'For N/A
            For Q = 1 To xMax
                If NAflag(Q) > 0 Then
                    exSheet.Cells(I, 7 + NAflag(Q)) = "N/A"
                    Y = Y + 1
                End If
            Next Q
            
            SQLQ = "SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJobCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            Do While Not rsTemp.EOF
                If Not IsNull(rsTemp("PC_CRSCODE")) Then
                    If InStr(1, strCourseCodes, "," & Trim(rsTemp("PC_CRSCODE")) & ",") > 0 Then
                        J = CoJobCode(Trim(rsTemp("PC_CRSCODE")))
                        If NAflag(J) > 0 Then
                            exSheet.Cells(I, 7 + NAflag(J)) = ""
                            'Ticket #22600 - Musashi - Do not supress Red color/Not Good
                            'If glbCompSerial <> "S/N - 2288W" Then
                                exSheet.Cells(I, 7 + NAflag(J) + 124) = "Not Good"
                            'End If
                        End If
                        xNA = xNA + 1
                    End If
                End If
                rsTemp.MoveNext
            Loop
            rsTemp.Close
        End If
        
        If Not IsNull(RsEdEmp("ES_CRSCODE")) Then
            If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
                'Green:   Training in Good Standing
                'Required Course with no renewal date
                'Required Course with a renewal date greater than 30 days from today
                'A non-required course with no renewal date
                '
                'Yellow:    Training will Expire in Thirty Days
                'Any course with a renewal date within 30 days of today
                '
                'Red:   Training Not in Good Standing
                'Required Course with no completed date.
                'Required Course not taken (No Continuing Education record)
                'Any course with a renewal date less than today’s date
                    
                'Check If Required Courses
                If IsDate(RsEdEmp("ES_DATCOMP")) Then
                    'Ticket #24875 - Day and Month switching in Excel if Day <= 12
                    If UCase(NewDateFormat) <> "MM/DD/YYYY" And UCase(NewDateFormat) <> "DD/MM/YYYY" Then
                        exSheet.Cells(I, 7 + NAflag(J)) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                    Else
                        exSheet.Cells(I, 7 + NAflag(J)) = Format(RsEdEmp("ES_DATCOMP"), "mm/dd/yyyy")
                    End If
                Else
                    exSheet.Cells(I, 7 + NAflag(J)) = ""
                End If
                strDisp = "Good"
                If InStr(1, strReqCourses, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                    flgReqC = True
                Else
                    flgReqC = False
                End If
                'for Good Standing - Green
                If flgReqC And Not IsDate(RsEdEmp("ES_RENEW")) Then
                    strDisp = "Good"
                End If
                If flgReqC And IsDate(RsEdEmp("ES_RENEW")) Then
                    If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) > 30 Then
                        strDisp = "Good"
                    End If
                End If
                If Not flgReqC And Not IsDate(RsEdEmp("ES_RENEW")) Then
                    strDisp = "Good"
                End If
                'Yellow:    Training will Expire in Thirty Days
                If IsDate(RsEdEmp("ES_RENEW")) Then
                    If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) >= 0 And DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) <= 30 Then
                        strDisp = "Expire"
                    End If
                End If
                'Red:   Training Not in Good Standing
                If IsDate(RsEdEmp("ES_RENEW")) Then
                    If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) < 0 Then
                        strDisp = "Not Good"
                    End If
                End If
                If flgReqC And Not IsDate(RsEdEmp("ES_DATCOMP")) Then
                    strDisp = "Not Good"
                End If
                'Hemu
                'Ticket #22600 - Musashi - do not suppress Red color/Not Good
                'If glbCompSerial <> "S/N - 2288W" Then
                    exSheet.Cells(I, 7 + NAflag(J) + 124) = strDisp     'Original
                'ElseIf strDisp <> "Not Good" Then
                '    exSheet.Cells(I, 7 + NAflag(J) + 124) = strDisp
                'End If
                'Hemu
            End If
        End If
NextLine02:
        K = K + 1
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    
    
    'Save new Excel file as XLS
    'exBook.SaveAs "C:\TrainMat.xls"
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    
    'exApp.Quit
    exApp.Visible = True
    
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
    Exit Sub
Err_XLS:
'    If Err.Number = 91 Then
'        MsgBox Err.Number
'        Resume Next
'    End If
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 1004 Then
        Resume Next
    End If
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Err = 76 Then
        MsgBox Err.Description & " to save the Training Matrix Report." & vbCrLf & "Please check Company Preference under Setup Menu."
        Exit Sub
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriter", "", "Select")
Resume Next
End Sub

Private Function getWSQLQ(WithCourse As Boolean)
Dim QStr
QStr = glbSeleDeptUn

If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
If clpDept <> "" Then QStr = QStr & " AND ED_DEPTNO in ('" & Replace(clpDept, ",", "','") & "')"
If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"
If clpCode(1) <> "" Then
    'QStr = QStr & " AND ED_ORG='" & clpCode(1) & "'"
    QStr = QStr & " AND ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
End If
If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
If clpPT <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpPT, ",", "','") & "')"
If clpProv <> "" Then QStr = QStr & " AND ED_PROV='" & clpProv & "'"
If clpProvEmp <> "" Then QStr = QStr & " AND ED_PROVEMP='" & clpProv & "'"
If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    If chkExclCONP.Visible And chkExclRET.Visible = True Then
        If chkExclCONP Then
            QStr = QStr & " AND ED_EMP <> 'CONP'"
        End If
        If chkExclRET Then
            QStr = QStr & " AND ED_EMP <> 'RET'"
        End If
    End If
End If

If WithCourse Then
    If IsDate(dlpDateRange(0)) Then QStr = QStr & " AND ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
    If IsDate(dlpDateRange(1)) Then QStr = QStr & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
    If clpCrsType <> "" Then QStr = QStr & " AND ES_CTYPE IN ('" & Replace(clpCrsType, ",", "','") & "') "
    If clpCrsCode <> "" Then QStr = QStr & " AND ES_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
    If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
        QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
        If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
        If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
        If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
        QStr = QStr & ")"
    End If
End If

getWSQLQ = QStr

End Function

Private Function getWSQLQ_TrainList(WithCourse As Boolean)
Dim QStr
QStr = glbSeleDeptUn

If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
If clpDept <> "" Then QStr = QStr & " AND ED_DEPTNO in ('" & Replace(clpDept, ",", "','") & "')"
If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"
If clpCode(1) <> "" Then
    'QStr = QStr & " AND ED_ORG='" & clpCode(1) & "'"
    QStr = QStr & " AND ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
End If
If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
If clpPT <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpPT, ",", "','") & "')"
If clpProv <> "" Then QStr = QStr & " AND ED_PROV='" & clpProv & "'"
If clpProvEmp <> "" Then QStr = QStr & " AND ED_PROVEMP='" & clpProv & "'"
If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "

'Ticket #29660 - Contract Employees Enhancement
If glbWFC Then
    If chkExclCONP.Visible And chkExclRET.Visible = True Then
        If chkExclCONP Then
            QStr = QStr & " AND ED_EMP <> 'CONP'"
        End If
        If chkExclRET Then
            QStr = QStr & " AND ED_EMP <> 'RET'"
        End If
    End If
End If

If WithCourse Then
    If IsDate(dlpDateRange(0)) Then QStr = QStr & " AND TR_COURSE_TAKEN >=" & Date_SQL(dlpDateRange(0).Text)
    If IsDate(dlpDateRange(1)) Then QStr = QStr & " AND TR_COURSE_TAKEN <=" & Date_SQL(dlpDateRange(1).Text)
    
    'Ticket #22274 - Using Course Code Master screen to filter by Course Type. No Course Type in Training List.
    'If clpCrsType <> "" Then QStr = QStr & " AND ES_CTYPE IN ('" & Replace(clpCrsType, ",", "','") & "') "
    If clpCrsType <> "" Then QStr = QStr & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_CTYPE IN ('" & Replace(clpCrsType, ",", "','") & "')) "
    
    If clpCrsCode <> "" Then QStr = QStr & " AND TR_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
    If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
        QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
        If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
        If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
        If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
        QStr = QStr & ")"
    End If
End If

getWSQLQ_TrainList = QStr
End Function

Private Sub AddEmpWrokList(xEmpNo)
Dim rslocWrk As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND TT_EMPNBR = " & xEmpNo & " "
    rslocWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rslocWrk.EOF Then
        rslocWrk.AddNew
        rslocWrk("TT_COMPNO") = "001"
        rslocWrk("TT_EMPNBR") = xEmpNo
        rslocWrk("TT_WRKEMP") = glbUserID
        rslocWrk.Update
    End If
    rslocWrk.Close
End Sub

Private Sub ScreenSetup4TrainingPlan()
    Label3.Visible = False
    Label4.Visible = False
    clpProv.Visible = False
    clpProvEmp.Visible = False
    chkShowEmp.Visible = False
    
    chkReqCourses.Visible = False
    chkLegisCourses.Visible = False
    chkCoursesTaken.Visible = False
    chkRenewalExce.Visible = False
    lblBCode(1).Visible = False
    clpCrsType.Visible = False
    
    lblShift.Top = lblBCode(0).Top
    txtShift.Top = clpCrsCode.Top
    lblBCode(0).Top = lblBCode(1).Top
    clpCrsCode.Top = clpCrsType.Top
    
    chkShowAllEmp.Top = 5620
    fraGroup.Left = 120
    fraGroup.Top = 6420
    fraGroup.Visible = True
    
    Call comGrpLoad
End Sub

Private Sub comGrpLoad()


    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")

    'If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "Course" ' Code"
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Course" ' Code"
    comGroup(1).AddItem "(none)"
    comGroup(2).AddItem "Course" ' Code"
    comGroup(2).AddItem "(none)"
    comGroup(3).AddItem "Renewal Date"
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 1
    comGroup(3).ListIndex = 0
    comGroup(3).Enabled = False

End Sub

Private Function Cri_SetAl_TrainPlan()
Dim x%
Dim SQLQ
Dim SQLR
Dim rsES As New ADODB.Recordset
Dim I As Integer
On Error GoTo modSetCriteria_Err

Cri_SetAl_TrainPlan = False

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
'DateSelCri = ""

Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere

For I = 0 To 5
    Call Cri_Code(I)
Next

Call Cri_FTDates
Call Cri_PT
Call Cri_EE
Call Cri_Job
Call Cri_PosGroup
Call Cri_CourseCode
Call Cri_Shift
'-------------

' report name
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rztrainplan.rpt"

x% = Cri_Sorts()   ' returns number of sections formated

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Training Plan Report"

Cri_SetAl_TrainPlan = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Training Plan", "Training Plan Report", "Select")
Cri_SetAl_TrainPlan = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String, fld$
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    CodeCri = ""
    If intIdx% = 0 Then fld$ = "HREMP.ED_LOC"
    If intIdx% = 1 Then fld$ = "HREMP.ED_ORG"
    If intIdx% = 2 Then fld$ = "HREMP.ED_EMP"
    If intIdx% = 3 Then fld$ = "HREMP.ED_REGION"
    If intIdx% = 4 Then fld$ = "HREMP.ED_ADMINBY"
    If intIdx% = 5 Then fld$ = "HREMP.ED_SECTION"

    If (intIdx% <> 0) Then
        CodeCri = "({" & fld$ & "} IN ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    Else
        CodeCri = "({" & fld$ & "} = '" & clpCode(intIdx%).Text & "')"
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

Private Sub Cri_PosGroup()
Dim EECri As String, OneSet%, x%
Dim strCa$

If Len(clpPosGroup.Text) < 1 Then Exit Sub

'EECri = "{HRJOB.ED_PT}= '" & clpJOB.Text & "'"
strCa$ = "HRJOB.JB_GRPCD"

If Len(clpPosGroup.Text) > 0 Then
    EECri = "({" & strCa$ & "} in ['" & Replace(clpPosGroup.Text, ",", "','") & "']) "  ' AND ({HR_JOB_HISTORY.JH_CURRENT}) "
End If


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub
Private Sub Cri_Job()
Dim EECri As String, OneSet%, x%
Dim strCa$

If Len(clpJob.Text) < 1 Then Exit Sub

strCa$ = "HR_JOB_HISTORY.JH_JOB"

If Len(clpJob.Text) > 0 Then
    EECri = "({" & strCa$ & "} in ['" & Replace(clpJob.Text, ",", "','") & "']) "
End If


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

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
Dim grpCond$, grpField$, I
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
' first set primary grouping
Y% = 0

grpField$ = getEGroup(comGroup(0).Text)
'If grpField$ <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    If comGroup(0).Text = "Course" Then
        grpField$ = "{tblCourseCode.TB_DESC}"
    End If
    If comGroup(0).Text = "(none)" Then
        grpField$ = "{HR_TRAIN.TR_COMPNO}"
    End If
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    strSFormat$ = "GH1;T;X;X;X;X;X;X"
    ''Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF1;T;X;X;X;X;X;X"
    ''Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{tblCourseCode.TB_DESC}"
        Case 2: grpField$ = "{HR_TRAIN.TR_COMPNO}" ' "(none)"
    End Select

    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$

    GrpIdx% = comGroup(2).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{tblCourseCode.TB_DESC}"
        Case 1: grpField$ = "{HR_TRAIN.TR_COMPNO}" '"(none)"
    End Select
    grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(2) = grpCond$


Cri_Sorts = z% ' next section number to format

End Function

Private Sub XLSwriter_All_TrainList()
Dim CoJobCodeS As New Collection
Dim CoJobCode As New Collection
Dim rsJobCode As New ADODB.Recordset
Dim RsEdEmp As New ADODB.Recordset
Dim rsCurJob As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim strCourseCodes As String
Dim SQLQ, I, J, K, L, M, x, Y, Q, xMax As Integer, xRecNum As Long
Dim exApp As Object, exBook As Object, exSheet As Object
Dim xlsFileTmp, xlsFileMat, xEmpnbr, xNA, xJobDesc, xJobCode, xLocation
Dim StartLine As Long, strTemp As String
Dim NewDateFormat
Dim CRSDesc(260)
Dim NAflag(260)
Dim flgReqC As Boolean, strDisp As String  'strReqCourses As String,

Dim z As Integer
Dim strCourse As String
Dim strAndWhere As String
Dim QStr, QStr1
Dim xRes

On Error GoTo Err_XLS

    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainTmp.xls"
    'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "TrainMat.xls"
    'Ticket# 8293
    If glbLinamar Then 'Or glbCompSerial = "S/N - 2336W" Then
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Reports\TrainMat" & Trim(glbUserID) & ".xls"
    Else
        xlsFileMat = xTrainMatrixPath & IIf(Right(xTrainMatrixPath, 1) = "\", "", "\") & "TrainMat" & Trim(glbUserID) & ".xls"
    End If

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Sub
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."

    FileCopy xlsFileTmp, xlsFileMat
    
    'Make a list of Required Courses....
    SQLQ = "SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE 1=1 "
    SQLQ = SQLQ & " AND PC_JOB IN (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & getWSQLQ_TrainList(False) & "))"
    If clpJob <> "" Then SQLQ = SQLQ & " AND PC_JOB = '" & clpJob.Text & "' "
    If clpCrsCode <> "" Then SQLQ = SQLQ & " AND PC_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
    If chkReqCourses Then If chkLegisCourses Then SQLQ = SQLQ & " AND PC_LEGISLATED <> 0"

    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    strReqCourses = ","
    Do While Not rsTemp.EOF
        strReqCourses = strReqCourses & Trim(rsTemp("PC_CRSCODE")) & ","
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    'End - List of Required Courses
    
    'Get list of Courses from HRTABL
    'Hemu - Since if there are more than 260 courses in the table it gives the Subscript out of range error
    '       and also Excel cannot take more than 124cols, the user may select only the courses
    '       needed in the report, made the following change to only select from the HRTABL the
    '       courses based on the selection criteria given by the user
    QStr = ""
    QStr1 = ""
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ESCD' "
    'If clpCrsCode <> "" Then SQLQ = SQLQ & " AND TB_KEY IN ('" & Replace(clpCrsCode, ",", "','") & "') "
    
    'And courses which are in Continuing Education screen only (when Required Courses NOT CHECKED)
    If (chkReqCourses.Value = 0) Then
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'SQLQ = SQLQ & " AND TB_KEY IN (SELECT ES_CRSCODE FROM HREDSEM WHERE ES_EMPNBR IN"
        SQLQ = SQLQ & " AND TB_KEY IN (SELECT TR_CRSCODE FROM HR_TRAIN WHERE TR_EMPNBR IN"
        SQLQ = SQLQ & " (SELECT ED_EMPNBR FROM HREMP "
        
        Call glbCri_DeptUN(clpDept.Text)
    
        QStr = Replace(Replace(glbstrSelCri, "{", "("), "}", ")")
        If glbSQL Or glbOracle Then
            QStr = Replace(Replace(QStr, "[", "("), "]", ")")
        End If
        If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
        If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"
        If clpCode(1) <> "" Then
            'QStr = QStr & " AND ED_ORG='" & clpCode(1) & "'"
            QStr = QStr & " AND ED_ORG in  ('" & Replace(clpCode(1).Text, ",", "','") & "')"
        End If
        If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
        If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
        If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
        If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
        If clpPT <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpPT, ",", "','") & "')"
        If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
            
        If clpProv <> "" Then QStr = QStr & " AND ED_PROV='" & clpProv & "'"
        If clpProvEmp <> "" Then QStr = QStr & " AND ED_PROVEMP='" & clpProv & "'"
            
        'Ticket #29660 - Contract Employees Enhancement
        If glbWFC Then
            If chkExclCONP.Visible And chkExclRET.Visible = True Then
                If chkExclCONP Then
                    QStr = QStr & " AND ED_EMP <> 'CONP'"
                End If
                If chkExclRET Then
                    QStr = QStr & " AND ED_EMP <> 'RET'"
                End If
            End If
        End If
        
        If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
            QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
            If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
            If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
            If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
            QStr = QStr & ")"
        End If
        
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'If IsDate(dlpDateRange(0)) Then QStr1 = QStr1 & " AND ES_DATCOMP>=" & Date_SQL(dlpDateRange(0).Text)
        If IsDate(dlpDateRange(0)) Then QStr1 = QStr1 & " AND TR_COURSE_TAKEN >=" & Date_SQL(dlpDateRange(0).Text)
        'If IsDate(dlpDateRange(1)) Then QStr1 = QStr1 & " AND ES_DATCOMP<=" & Date_SQL(dlpDateRange(1).Text)
        If IsDate(dlpDateRange(1)) Then QStr1 = QStr1 & " AND TR_COURSE_TAKEN <=" & Date_SQL(dlpDateRange(1).Text)
        
        'Ticket #22274 - City of Chatham-Kent - No Course Type in Training List so using the Course Code Master.
        'If clpCrsType <> "" Then QStr1 = QStr1 & " AND ES_CTYPE IN ('" & Replace(clpCrsType, ",", "','") & "') "
        If clpCrsType <> "" Then QStr1 = QStr1 & " AND TR_CRSCODE IN (SELECT ES_CRSCODE FROM HR_COURSECODE_MASTER WHERE ES_CTYPE IN ('" & Replace(clpCrsType, ",", "','") & "') )"
        
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'If clpCrsCode <> "" Then QStr1 = QStr1 & " AND ES_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
        If clpCrsCode <> "" Then QStr1 = QStr1 & " AND TR_CRSCODE IN ('" & Replace(clpCrsCode, ",", "','") & "') "
        QStr = QStr & ")" & QStr1
        
        SQLQ = SQLQ & " WHERE " & QStr
        SQLQ = SQLQ & ")"
    End If
    'End - Courses in the Continuing Education screen only (when Required Courses NOT CHECKED)
    
    SQLQ = SQLQ & " ORDER BY TB_KEY"
    rsJobCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    I = 1: strCourseCodes = ","
    
    'Create a separate collection for courses which are Required
    Do While Not rsJobCode.EOF
        If (chkReqCourses) Then
            'Collection containing Required Courses only
            If InStr(strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
                CoJobCodeS.Add Trim(rsJobCode("TB_KEY"))
                CoJobCode.Add I, Trim(rsJobCode("TB_KEY"))
                CRSDesc(I) = Trim(rsJobCode("TB_DESC"))
            End If
        Else
            'All Courses from Continuing Education screen
            CoJobCodeS.Add Trim(rsJobCode("TB_KEY"))
            CoJobCode.Add I, Trim(rsJobCode("TB_KEY"))
            CRSDesc(I) = Trim(rsJobCode("TB_DESC"))
        End If
        
        'Set value to NAflag(x) to indicate if course is Required Course (1) or not (0).
        'Independent of the collection created just above
        If InStr(strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
            NAflag(I) = 1
        Else
            NAflag(I) = 0
        End If

        'Create a string of courses (required or not) and count them
        If (chkReqCourses) Then
            If InStr(1, strReqCourses, "," & Trim(rsJobCode("TB_KEY")) & ",") > 0 Then
                strCourseCodes = strCourseCodes & Trim(rsJobCode("TB_KEY")) & ","
                I = I + 1
            End If
        Else
            strCourseCodes = strCourseCodes & Trim(rsJobCode("TB_KEY")) & ","
            I = I + 1
        End If
        rsJobCode.MoveNext
    Loop
    xMax = CoJobCode.count
    rsJobCode.Close
    'End - List/collection/string of courses created from HRTABL and/or Continuing Education screen
    
    'Check if all the courses can fit into the Excel spreadsheet, warn if necessary.
    If xMax > 124 Then
        xRes = MsgBox("Course Code exceeds 124 courses: Training Matrix report will not fit in MS Excel spreadsheet." & Chr(10) & "Use the Selection Criteria to narrow down the courses." & Chr(10) & Chr(10) & "Exiting Training Matrix report.", vbOKOnly, "info:HR - Course Code exceeds 124 columns")
        Exit Sub
    End If
    
    'To get which columns/cells are all N/A
    If Not glbOracle Then
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'SQLQ = "SELECT ES_CRSCODE "
        'SQLQ = SQLQ & " FROM HREDSEM INNER JOIN HREMP ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
        'SQLQ = SQLQ & " WHERE ES_CRSCODE IS NOT NULL "
        'SQLQ = SQLQ & " AND " & getWSQLQ(True) & " "
        SQLQ = "SELECT TR_CRSCODE "
        SQLQ = SQLQ & " FROM HR_TRAIN INNER JOIN HREMP ON HR_TRAIN.TR_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE TR_CRSCODE IS NOT NULL "
        SQLQ = SQLQ & " AND " & getWSQLQ_TrainList(True) & " "
    Else
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'SQLQ = "SELECT ES_CRSCODE "
        'SQLQ = SQLQ & " FROM HREDSEM,HREMP "
        'SQLQ = SQLQ & " WHERE  ES_CRSCODE IS NOT NULL "
        'SQLQ = SQLQ & " AND " & getWSQLQ(True) & " "
        'SQLQ = SQLQ & " AND (HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR)"
        SQLQ = "SELECT TR_CRSCODE "
        SQLQ = SQLQ & " FROM HR_TRAIN,HREMP "
        SQLQ = SQLQ & " WHERE TR_CRSCODE IS NOT NULL "
        SQLQ = SQLQ & " AND " & getWSQLQ_TrainList(True) & " "
        SQLQ = SQLQ & " AND (HR_TRAIN.TR_EMPNBR = HREMP.ED_EMPNBR)"
    End If
    'Ticket #22274 - City of Chatham-Kent - changing to Training List
    'SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HREDSEM.ES_DATCOMP "
    SQLQ = SQLQ & "ORDER BY HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_SURNAME,HREMP.ED_FNAME,HR_TRAIN.TR_COURSE_TAKEN "
    
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xRecNum = 0
    xRecNum = RsEdEmp.RecordCount
    K = 0: I = StartLine + 1: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = "": xLocation = ""
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
        If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("TR_CRSCODE")) & ",") > 0 Then
            'Ticket #22274 - City of Chatham-Kent - changing to Training List
           'J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
           J = CoJobCode(Trim(RsEdEmp("TR_CRSCODE")))
           NAflag(J) = 1
        End If
NextLine01:
        K = K + 1
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    
    Y = 1
    For x = 1 To xMax
        If NAflag(x) > 0 Then 'How many course codes display and on which field in xls
            NAflag(x) = Y
            Y = Y + 1
        End If
    Next x
    'End - To get whick columns are all N/A
    
    'Create new WorkBook of Excel
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    Set exSheet = exBook.Worksheets(1)

    
    exSheet.Cells(1, 5) = "Training Matrix"
    exSheet.Cells(1, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
    exSheet.Cells(2, 1) = "Time: " & Time$
    If Not (IsDate(dlpDateRange(0).Text) Or IsDate(dlpDateRange(1).Text)) Then
        exSheet.Cells(2, 5) = "No date entered"
    Else
        strTemp = ""
        If IsDate(dlpDateRange(0).Text) Then
            strTemp = "From Date: " & Format(dlpDateRange(0).Text, "mmm dd, yyyy") & "  "
        End If
        If IsDate(dlpDateRange(1).Text) Then
            strTemp = strTemp & "To Date: " & Format(dlpDateRange(1).Text, "mmm dd, yyyy")
        End If
        exSheet.Cells(2, 5) = strTemp
    End If
    'Ticket 4131 Frank 05/27/2003
    If glbWFC Then
        StartLine = 6
    Else
        StartLine = 8
    End If
    
    exSheet.Cells(StartLine, 1) = lStr("Division")
    exSheet.Cells(StartLine, 2) = lStr("Department")
    exSheet.Cells(StartLine, 3) = "Employee #"
    exSheet.Cells(StartLine, 4) = "Name"
    exSheet.Cells(StartLine, 5) = "Job Title"
    'Zahoor(Sam) 03/02/2006
    exSheet.Cells(StartLine, 6) = "Location"
    'Zahoor(Sam) 03/02/2006
    exSheet.Cells(StartLine, 7) = lStr("Original Hire")
    exSheet.Cells(StartLine, 8) = "Course Taken"
    
    
    StartLine = StartLine + 1
    Y = 1
    'Display Course Codes and Descriptions on Title
    For I = 1 To CoJobCode.count
        If NAflag(I) > 0 Then
            exSheet.Cells(StartLine, 7 + Y) = CoJobCodeS.Item(I)
            exSheet.Cells(StartLine + 1, 7 + Y) = CRSDesc(I)

            SQLQ = "SELECT HR_JOB_COURSE.PC_CRSCODE, HR_JOB_COURSE.PC_JOB, HR_JOB_COURSE.PC_LEGISLATED FROM HR_JOB_COURSE "
            SQLQ = SQLQ & "WHERE HR_JOB_COURSE.PC_CRSCODE = '" & CoJobCodeS.Item(I) & "' "
            SQLQ = SQLQ & " AND PC_JOB IN (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & getWSQLQ_TrainList(False) & "))"
            If clpJob <> "" Then SQLQ = SQLQ & " AND PC_JOB = '" & clpJob.Text & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                If rsTemp("PC_LEGISLATED") Then
                    exSheet.Cells(StartLine + 2, 7 + Y) = "R/L"
                Else
                    exSheet.Cells(StartLine + 2, 7 + Y) = "R"
                End If
            End If
            rsTemp.Close
            Y = Y + 1
        End If
    Next
    'Ticket #22274 - City of Chatham-Kent - changing to Training List
    'SQLQ = "SELECT ES_CRSCODE, ES_DATCOMP, ES_RENEW, "
    SQLQ = "SELECT TR_CRSCODE, TR_COURSE_TAKEN, TR_RENEW, "
    
    'new query added ED_LOC on the request of City of Chatham-Kent -tkt #2188 request
    SQLQ = SQLQ & " ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DIV, ED_DEPTNO,ED_DOH,ED_LOC "
    'Zahoor(Sam) 03/02/2006
    
    If Not glbOracle Then
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'SQLQ = SQLQ & " FROM HREMP LEFT JOIN HREDSEM ON HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " FROM HREMP LEFT JOIN HR_TRAIN ON HR_TRAIN.TR_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE "
    Else
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'SQLQ = SQLQ & " FROM HREDSEM,HREMP "
        SQLQ = SQLQ & " FROM HR_TRAIN,HREMP "
        'Ticket #15688 - Begin 'wrong direction of left join for Oracle
        'SQLQ = SQLQ & " WHERE (HREDSEM.ES_EMPNBR = HREMP.ED_EMPNBR(+))"
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'SQLQ = SQLQ & " WHERE (HREMP.ED_EMPNBR = HREDSEM.ES_EMPNBR(+))"
        SQLQ = SQLQ & " WHERE (HREMP.ED_EMPNBR = HR_TRAIN.TR_EMPNBR(+))"
        'Ticket #15688 - End
        SQLQ = SQLQ & " AND "
    End If
    If chkShowEmp Then
        SQLQ = SQLQ & getWSQLQ_TrainList(False) & " "
    Else
        SQLQ = SQLQ & getWSQLQ_TrainList(True) & " "
    End If
    'Ticket #22274 - City of Chatham-Kent - changing to Training List
    'SQLQ = SQLQ & "ORDER BY ED_DIV,ED_DEPTNO,ED_SURNAME,ED_FNAME,ES_DATCOMP "
    SQLQ = SQLQ & "ORDER BY ED_DIV,ED_DEPTNO,ED_SURNAME,ED_FNAME,TR_COURSE_TAKEN "
    
    If RsEdEmp.State <> 0 Then RsEdEmp.Close
    RsEdEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    xRecNum = RsEdEmp.RecordCount
    
    K = 0: I = StartLine + 2: xEmpnbr = -1: xNA = 0: xJobDesc = "": xJobCode = "": xLocation = ""
    Y = 1
    NewDateFormat = UCase(glbsDateFormat)
    If InStr(1, NewDateFormat, "YYYY") = 0 Then
        NewDateFormat = Replace(NewDateFormat, "YY", "YYYY")
    End If
    
    Do While Not RsEdEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xRecNum) * 100
        If RsEdEmp("ED_EMPNBR") <> xEmpnbr Then
            xEmpnbr = RsEdEmp("ED_EMPNBR")

            I = I + 1: xNA = 0
            
            SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB "
            SQLQ = SQLQ & " WHERE JB_CODE IN (SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR = " & xEmpnbr & ")"
            rsCurJob.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            xJobDesc = "": xJobCode = ""
            If Not rsCurJob.EOF Then
                If Not IsNull(rsCurJob("JB_DESCR")) Then
                    xJobDesc = rsCurJob("JB_DESCR")
                    xJobCode = rsCurJob("JB_CODE")
                End If
            End If
            rsCurJob.Close
            
            exSheet.Cells(I, 1) = RsEdEmp("ED_DIV")
            exSheet.Cells(I, 2) = RsEdEmp("ED_DEPTNO")
            exSheet.Cells(I, 3) = RsEdEmp("ED_EMPNBR")
            exSheet.Cells(I, 4) = RsEdEmp("ED_SURNAME") & "," & RsEdEmp("ED_FNAME")
            exSheet.Cells(I, 5) = xJobDesc '
           ''Zahoor(Sam) 03/02/2006
            exSheet.Cells(I, 6) = RsEdEmp("ED_LOC")
            'Zahoor(Sam) 03/02/2006
            exSheet.Cells(I, 7) = Format(RsEdEmp("ED_DOH"), NewDateFormat) '"SHORT DATE")
           
            
            'For N/A
            For Q = 1 To xMax
                If NAflag(Q) > 0 Then
                    exSheet.Cells(I, 7 + NAflag(Q)) = "N/A"
                    Y = Y + 1
                End If
            Next Q
            
            SQLQ = "SELECT PC_CRSCODE FROM HR_JOB_COURSE WHERE PC_JOB = '" & xJobCode & "' "
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            Do While Not rsTemp.EOF
                If Not IsNull(rsTemp("PC_CRSCODE")) Then
                    If InStr(1, strCourseCodes, "," & Trim(rsTemp("PC_CRSCODE")) & ",") > 0 Then
                        J = CoJobCode(Trim(rsTemp("PC_CRSCODE")))
                        If NAflag(J) > 0 Then
                            exSheet.Cells(I, 7 + NAflag(J)) = ""
                            'Ticket #22600 - Musashi - Do not suppress Red color/Not Good
                            'If glbCompSerial <> "S/N - 2288W" Then
                                exSheet.Cells(I, 7 + NAflag(J) + 124) = "Not Good"
                            'End If
                        End If
                        xNA = xNA + 1
                    End If
                End If
                rsTemp.MoveNext
            Loop
            rsTemp.Close
        End If
        
        'Ticket #22274 - City of Chatham-Kent - changing to Training List
        'If Not IsNull(RsEdEmp("ES_CRSCODE")) Then
        If Not IsNull(RsEdEmp("TR_CRSCODE")) Then
            'Ticket #22274 - City of Chatham-Kent - changing to Training List
            'If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
            If InStr(1, strCourseCodes, "," & Trim(RsEdEmp("TR_CRSCODE")) & ",") > 0 Then
                'J = CoJobCode(Trim(RsEdEmp("ES_CRSCODE")))
                J = CoJobCode(Trim(RsEdEmp("TR_CRSCODE")))
                
                'Green:   Training in Good Standing
                'Required Course with no renewal date
                'Required Course with a renewal date greater than 30 days from today
                'A non-required course with no renewal date
                '
                'Yellow:    Training will Expire in Thirty Days
                'Any course with a renewal date within 30 days of today
                '
                'Red:   Training Not in Good Standing
                'Required Course with no completed date.
                'Required Course not taken (No Continuing Education record)
                'Any course with a renewal date less than today’s date
                    
                'Check If Required Courses
                'Ticket #22274 - City of Chatham-Kent - changing to Training List
                'If IsDate(RsEdEmp("ES_DATCOMP")) Then
                If IsDate(RsEdEmp("TR_COURSE_TAKEN")) Then
                    'exSheet.Cells(I, 7 + NAflag(J)) = Format(RsEdEmp("ES_DATCOMP"), NewDateFormat)
                    exSheet.Cells(I, 7 + NAflag(J)) = Format(RsEdEmp("TR_COURSE_TAKEN"), NewDateFormat)
                Else
                    exSheet.Cells(I, 7 + NAflag(J)) = ""
                End If
                
                strDisp = "Good"
                
                'Ticket #22274 - City of Chatham-Kent - changing to Training List
                'If InStr(1, strReqCourses, "," & Trim(RsEdEmp("ES_CRSCODE")) & ",") > 0 Then
                If InStr(1, strReqCourses, "," & Trim(RsEdEmp("TR_CRSCODE")) & ",") > 0 Then
                    flgReqC = True
                Else
                    flgReqC = False
                End If
                
                'for Good Standing - Green
                'Ticket #22274 - City of Chatham-Kent - changing to Training List
                'If flgReqC And Not IsDate(RsEdEmp("ES_RENEW")) Then
                If flgReqC And Not IsDate(RsEdEmp("TR_RENEW")) Then
                    strDisp = "Good"
                End If
                'If flgReqC And IsDate(RsEdEmp("ES_RENEW")) Then
                If flgReqC And IsDate(RsEdEmp("TR_RENEW")) Then
                    'If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) > 30 Then
                    If DateDiff("d", Now, CVDate(RsEdEmp("TR_RENEW"))) > 30 Then
                        strDisp = "Good"
                    End If
                End If
                'If Not flgReqC And Not IsDate(RsEdEmp("ES_RENEW")) Then
                If Not flgReqC And Not IsDate(RsEdEmp("TR_RENEW")) Then
                    strDisp = "Good"
                End If
                
                'Yellow:    Training will Expire in Thirty Days
                'Ticket #22274 - City of Chatham-Kent - changing to Training List
                'If IsDate(RsEdEmp("ES_RENEW")) Then
                If IsDate(RsEdEmp("TR_RENEW")) Then
                    'If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) >= 0 And DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) <= 30 Then
                    If DateDiff("d", Now, CVDate(RsEdEmp("TR_RENEW"))) >= 0 And DateDiff("d", Now, CVDate(RsEdEmp("TR_RENEW"))) <= 30 Then
                        strDisp = "Expire"
                    End If
                End If
                
                'Red:   Training Not in Good Standing
                'Ticket #22274 - City of Chatham-Kent - changing to Training List
                'If IsDate(RsEdEmp("ES_RENEW")) Then
                If IsDate(RsEdEmp("TR_RENEW")) Then
                    'If DateDiff("d", Now, CVDate(RsEdEmp("ES_RENEW"))) < 0 Then
                    If DateDiff("d", Now, CVDate(RsEdEmp("TR_RENEW"))) < 0 Then
                        strDisp = "Not Good"
                    End If
                End If
                'If flgReqC And Not IsDate(RsEdEmp("ES_DATCOMP")) Then
                If flgReqC And Not IsDate(RsEdEmp("TR_COURSE_TAKEN")) Then
                    strDisp = "Not Good"
                End If
                
                'Hemu
                'Ticket #22600 - Musashi - Do not suppress Red color/Not Good
                'If glbCompSerial <> "S/N - 2288W" Then
                    exSheet.Cells(I, 7 + NAflag(J) + 124) = strDisp     'Original
                'ElseIf strDisp <> "Not Good" Then
                '    exSheet.Cells(I, 7 + NAflag(J) + 124) = strDisp
                'End If
                'Hemu
            End If
        End If
NextLine02:
        K = K + 1
        RsEdEmp.MoveNext
    Loop
    RsEdEmp.Close
    
    
    'Save new Excel file as XLS
    'exBook.SaveAs "C:\TrainMat.xls"
    exBook.Save
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Call Pause(1)
    'launch Excel file
    'Shell "Start " & GetShortName(xlsFileMat)
    If Not LanchXlsW98(xlsFileMat) Then
        Shell "cmd /c " & GetShortName(xlsFileMat)
    End If
    
    Exit Sub
Err_XLS:
'    If Err.Number = 91 Then
'        MsgBox Err.Number
'        Resume Next
'    End If
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    Screen.MousePointer = DEFAULT
    
    If Err = 1004 Then
        Resume Next
    End If
    
    If Err = 75 Then
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Not exBook Is Nothing Then
        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
    End If
    If Err = 70 Then
        Set exApp = Nothing
        MsgBox Err.Description & Chr(10) & "Please close Excel Files."
        Exit Sub
    End If
    If Err = 76 Then
        MsgBox Err.Description & " to save the Training Matrix Report." & vbCrLf & "Please check Company Preference under Setup Menu."
        Exit Sub
    End If
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "XLSwriter_TrainingList", "", "Select")
Resume Next
End Sub

Private Sub AddEmpToWrk(rsEMP As ADODB.Recordset, xType)
Dim SQLQ As String
Dim rsWRK As New ADODB.Recordset
If Not rsEMP.EOF Then
    SQLQ = "SELECT * FROM HREMPWRK WHERE TT_WRKEMP = '" & glbUserID & "' AND TT_EMPNBR = " & rsEMP("ED_EMPNBR")
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If rsWRK.EOF Then
        rsWRK.AddNew
        rsWRK("TT_EMPNBR") = rsEMP("ED_EMPNBR")
        rsWRK("TT_NEWDIV") = rsEMP("ED_DIV")
        rsWRK("TT_NEWDEPT") = rsEMP("ED_DEPTNO")
        rsWRK("TT_NAMEFLD") = Left(rsEMP("ED_SURNAME") & "," & rsEMP("ED_FNAME"), 40)
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK("TT_GRADE") = xType
        rsWRK.Update
    End If
End If
End Sub
