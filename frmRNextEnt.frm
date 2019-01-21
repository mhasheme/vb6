VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRNextEnt 
   Caption         =   "Future Entitlements Report"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9180
   ScaleWidth      =   10935
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   120
      TabIndex        =   25
      Top             =   0
      Width           =   9495
      Begin Threed.SSCheck chkAnnualized 
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
         _Version        =   65536
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Annualized Entitlement"
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
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2235
         MaxLength       =   4
         TabIndex        =   13
         Tag             =   "00-Employee Position Shift"
         Top             =   4290
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   2005
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Tag             =   "First Level of grouping records"
         Top             =   7800
         Width           =   2325
      End
      Begin VB.ComboBox comGroup 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   2005
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Tag             =   "First Level of grouping records"
         Top             =   7485
         Width           =   2325
      End
      Begin VB.ComboBox cmbHours 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRNextEnt.frx":0000
         Left            =   2220
         List            =   "frmRNextEnt.frx":0002
         TabIndex        =   9
         Tag             =   "Choose display in hours or days"
         Text            =   "Combo1"
         Top             =   2940
         Width           =   1215
      End
      Begin Threed.SSCheck chkEntitle 
         Height          =   225
         Left            =   120
         TabIndex        =   20
         Tag             =   "If X-Show Hourly Entitlements"
         Top             =   6120
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   2646
         _ExtentY        =   1323
         _StockProps     =   78
         Caption         =   "   Show Hourly Entitlements"
         ForeColor       =   0
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   5
         Tag             =   "00-Enter Status Code"
         Top             =   1590
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
         TabIndex        =   6
         Tag             =   "EDPT-Category"
         Top             =   1920
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
         TabIndex        =   4
         Tag             =   "00-Enter Union Code"
         Top             =   1260
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
         TabIndex        =   3
         Tag             =   "00-Enter Location Code"
         Top             =   930
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Tag             =   "00-Specific Department Desired"
         Top             =   600
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
         TabIndex        =   1
         Tag             =   "00-Specific Division Desired"
         Top             =   270
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
         Index           =   5
         Left            =   1920
         TabIndex        =   12
         Tag             =   "00-Enter Section Code"
         Top             =   3960
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
         Tag             =   "00-Enter Administered By Code"
         Top             =   3630
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
         Left            =   1920
         TabIndex        =   10
         Tag             =   "00-Enter Region Code"
         Top             =   3300
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Tag             =   "10-Enter Employee Number"
         Top             =   2250
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
         Left            =   3570
         TabIndex        =   16
         Tag             =   "40-Entitlement To Date Range"
         Top             =   4635
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   15
         Tag             =   "40-Entitlement From Date Range"
         Top             =   4635
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin Threed.SSCheck chkSick 
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Tag             =   "If X-Show Attendance Details"
         Top             =   5820
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Sick Time"
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
         Value           =   -1  'True
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   6
         Left            =   5280
         TabIndex        =   14
         Tag             =   "00-Enter Section Code"
         Top             =   4290
         Visible         =   0   'False
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   8
      End
      Begin INFOHR_Controls.EmployeeLookup elpReptau 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Tag             =   "10-Reporting Authority"
         Top             =   2595
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   503
         ShowUnassigned  =   1
         ShowDescription =   0   'False
         TextBoxWidth    =   7195
         RefreshDescriptionWhen=   2
         MultiSelect     =   -1  'True
      End
      Begin Threed.SSCheck chkVacation 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Tag             =   "If X-Show Attendance Details"
         Top             =   5520
         Width           =   2835
         _Version        =   65536
         _ExtentX        =   5001
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "   Show Vacation Time"
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
         Value           =   -1  'True
      End
      Begin INFOHR_Controls.DateLookup dlpEffDate 
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Tag             =   "40-Effective Date"
         Top             =   4965
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin VB.Label lblEffectDate 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date"
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
         Left            =   60
         TabIndex        =   45
         Top             =   5040
         Width           =   1245
      End
      Begin VB.Label lblRepAuth 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reporting Authority"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   44
         Top             =   2640
         Width           =   1350
      End
      Begin VB.Label lblSalDist 
         Caption         =   "Salary Distribution"
         Height          =   315
         Left            =   3420
         TabIndex        =   43
         Top             =   4290
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Future Entitlement"
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
         Left            =   60
         TabIndex        =   42
         Top             =   4680
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label lblShift 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Shift"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   41
         Top             =   4350
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
         Left            =   60
         TabIndex        =   40
         Top             =   1920
         Width           =   630
      End
      Begin VB.Label lblSection 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   39
         Top             =   3960
         Width           =   540
      End
      Begin VB.Label lblRegion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   38
         Top             =   3300
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
         Left            =   60
         TabIndex        =   37
         Top             =   3630
         Width           =   1125
      End
      Begin VB.Label lblLocation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   36
         Top             =   930
         Width           =   615
      End
      Begin VB.Label lblGrp 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Final Sort"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   35
         Top             =   7830
         Width           =   660
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
         Left            =   90
         TabIndex        =   34
         Top             =   7515
         Width           =   885
      End
      Begin VB.Label lblRepGrp 
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
         TabIndex        =   33
         Top             =   7155
         Width           =   1575
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
         TabIndex        =   32
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Display"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   31
         Top             =   3000
         Width           =   510
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   30
         Top             =   1605
         Width           =   450
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   29
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   28
         Top             =   2250
         Width           =   1290
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   60
         TabIndex        =   26
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   24
      Top             =   8475
      Width           =   9855
   End
   Begin VB.VScrollBar scrControl 
      Height          =   7995
      LargeChange     =   315
      Left            =   9720
      Max             =   100
      SmallChange     =   315
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   300
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9360
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
      MarginTop       =   720
      MarginBottom    =   720
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmRNextEnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private snapEntitle As New ADODB.Recordset
Private snapFuture As New ADODB.Recordset
Private snapHourEntitle As New ADODB.Recordset
Private rsEntRules As New ADODB.Recordset
Private fglbAsOf As String, fglbToDate As String
Private strGroup As String

Private Function getGroupField() As String

    Dim strTable As String
    Dim SQLQ As String
    
    Select Case comGroup(0).Text
        Case lStr("Division")
            strTable = ""
            strGroup = "HREMP.ED_DIV"
            SQLQ = "SELECT DISTINCT HREMP.ED_DIV as grpField, HR_DIVISION.Division_Name as TB_DESC FROM ((HREMP LEFT OUTER JOIN HR_DIVISION ON HREMP.ED_DIV = HR_DIVISION.DIV) "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Department")
            strTable = ""
            strGroup = "HREMP.ED_DEPTNO"
            SQLQ = "SELECT DISTINCT HREMP.ED_DEPTNO as grpField, HRDEPT.DF_NAME as TB_DESC FROM (HREMP INNER JOIN HRDEPT ON HREMP.ED_DEPTNO = HRDEPT.DF_NBR) INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
            SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Location")
            strGroup = "HREMP.ED_LOC"
            strTable = "HREMP.ED_LOC_TABL"
        Case lStr("Union")
            strGroup = "HREMP.ED_ORG"
            strTable = "HREMP.ED_ORG_TABL"
        Case lStr("Administered By")
            strGroup = "HREMP.ED_ADMINBY"
            strTable = "HREMP.ED_ADMINBY_TABL"
        Case "Employee Name"
            strTable = ""
            strGroup = "HREMP.ED_SURNAME"
            If glbOracle Then
                SQLQ = "SELECT DISTINCT HREMP.ED_SURNAME AS grpfield, CONCAT(CONCAT(HREMP.ED_SURNAME, ', '), HREMP.ED_FNAME) AS TB_DESC FROM HREMP "
            Else
                SQLQ = "SELECT DISTINCT HREMP.ED_SURNAME AS grpfield, ED_SURNAME + ', ' + ED_FNAME AS TB_DESC FROM HREMP "
            End If
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Section")
            strGroup = "HREMP.ED_SECTION"
            strTable = "HREMP.ED_SECTION_TABL"
        Case "Employment Type"
            strTable = ""
            strGroup = "HREMP.ED_EMPTYPE"
            SQLQ = "SELECT DISTINCT HREMP.ED_EMPTYPE as grpfield, ED_EMPTYPE as TB_DESC FROM HREMP "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case ("Home Line")
            strGroup = "HREMP.ED_HOMELINE"
            strTable = "HREMP.ED_HOMELINE_TABL"
        Case "Shift"
            strTable = ""
            strGroup = "HR_JOB_HISTORY.JH_SHIFT"
            SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_SHIFT as grpField, HR_JOB_HISTORY.JH_SHIFT as TB_DESC FROM HREMP LEFT OUTER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
        Case lStr("Rept. Authority 1")
            strTable = ""
            strGroup = "HR_JOB_HISTORY.JH_REPTAU"
            If glbOracle Then
                SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_REPTAU AS grpField, CONCAT(CONCAT(tblSuper.ED_SURNAME, ', '), tblSuper.ED_FNAME) AS TB_DESC "
                SQLQ = SQLQ & "FROM (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP tblSuper ON HR_JOB_HISTORY.JH_REPTAU = tblSuper.ED_EMPNBR "
                SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "

            Else
                SQLQ = "SELECT DISTINCT HR_JOB_HISTORY.JH_REPTAU AS grpField, tblSuper.ED_SURNAME + ', ' + tblSuper.ED_FNAME AS TB_DESC "
                SQLQ = SQLQ & "FROM (HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) INNER JOIN HREMP AS tblSuper ON HR_JOB_HISTORY.JH_REPTAU = tblSuper.ED_EMPNBR "
                SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
            End If
        Case lStr("Region")
            strGroup = "HREMP.ED_REGION"
            strTable = "HREMP.ED_REGION_TABL"
        Case "(none)"
            strGroup = "HREMP.ED_COMPNO"
            strTable = ""
            SQLQ = "SELECT DISTINCT " & strGroup & " as grpField, 'None' as TB_DESC FROM HREMP "
            SQLQ = SQLQ & "INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
    End Select
    
    If strTable <> "" Then
        SQLQ = "SELECT DISTINCT " & strGroup & " as grpField, HRTABL.TB_DESC as TB_DESC FROM ((HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        SQLQ = SQLQ & "LEFT OUTER JOIN HRTABL ON " & strTable & " = HRTABL.TB_NAME AND " & strGroup & " = HRTABL.TB_KEY) "
        SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) and "
    End If
    
    SQLQ = SQLQ & glbstrSelCri
    SQLQ = SQLQ & " ORDER BY " & strGroup
    
    If Not (glbSQL Or glbOracle) Then
        SQLQ = Replace(SQLQ, "INNER", "LEFT")
    End If
    
    getGroupField = SQLQ
End Function

Private Sub modAnnHours()
'laura 03/04/98

Dim dblServiceYears#, EmpNo As Long
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct#
Dim SQLQ As String, NumRec As Integer
Dim snapDuplic As New ADODB.Recordset
Dim oldEntitleUpd
Dim xKey, fglbWDate$, Accum  As Boolean

On Error GoTo modannhours_Err

Accum = False 'default the Accumulator uintil a decision is made


Screen.MousePointer = HOURGLASS

MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(0).FloodPercent = 0
pct# = 0

EmpNo = snapHourEntitle("ED_EMPNBR")

'if no records exist for this employee create one
    If CR_snapFuture(EmpNo, "HOUR") = 0 Then
        snapFuture.AddNew
        snapFuture("EN_COMPNO") = "001"
        snapFuture("EN_EMPNBR") = EmpNo&
        snapFuture("EN_ENTSORT") = 3
        snapFuture("EN_TYPE_TABL") = "ENTT"
        snapFuture("EN_TYPE") = rsEntRules("EH_HETYPE")
        snapFuture("EN_DESC") = rsEntRules("TB_DESC")
        snapFuture("EN_FDATE") = dlpDateRange(0).Text
        snapFuture("EN_TDATE") = dlpDateRange(1).Text
        snapFuture("EN_DHRS") = snapHourEntitle("JH_DHRS")
        snapFuture("EN_SURNAME") = snapHourEntitle("ED_SURNAME")
        snapFuture("EN_FNAME") = snapHourEntitle("ED_FNAME")
        snapFuture("EN_WRKEMP") = glbUserID
        snapFuture("EN_LDATE") = Now
        snapFuture("EN_LTIME") = Time$
        snapFuture("EN_LUSER") = glbUserID
        snapFuture.Update
    End If

BeginTrans

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select


    If IsNull(snapHourEntitle(fglbWDate$)) Then
        GoTo lblNext2Rec
    End If

    varStartDate = snapHourEntitle(fglbWDate$)  ' set start date
    If Not IsNumeric(snapHourEntitle("JH_DHRS")) Then
        dblDHours# = 0
    Else
        dblDHours# = snapHourEntitle("JH_DHRS")
    End If
    If Not IsNumeric(snapHourEntitle("JH_FTENUM")) Then
        dblFTEHours# = 0
    Else
        dblFTEHours# = snapHourEntitle("JH_FTENUM")
    End If

    'dblServiceYears# = MonthDiff(CVDate(varStartDate), Date)
    'dblServiceYears# = MonthDiff(CVDate(varStartDate), dlpDateRange(0).Text)
    dblServiceYears# = MonthDiff(CVDate(varStartDate), dlpEffDate.Text)
    If dblServiceYears# < 0 Then GoTo lblNext2Rec     'laura 03/06/98
    intWhereFit& = -1   ' first record can be just less than


    If IsNumeric(Val(rsEntRules("EH_BMONTH"))) And rsEntRules("EH_EMONTH") = "" Then
        If dblServiceYears# >= CDbl(Val(rsEntRules("EH_BMONTH"))) Then
            intWhereFit& = x%
        End If
    End If
    If IsNumeric(rsEntRules("EH_BMONTH")) And IsNumeric(rsEntRules("EH_EMONTH")) Then
        If dblServiceYears# >= CDbl(Val(rsEntRules("EH_BMONTH"))) And dblServiceYears# <= CDbl(Val(rsEntRules("EH_EMONTH"))) Then
            intWhereFit& = x%
        End If
    End If


    If intWhereFit& = -1 Then GoTo lblNext2Rec  ' skip record if not in any of the ranges

    dblNewEntitle# = Val(rsEntRules("EH_ENTITLE"))
    If rsEntRules("EH_TYPE") = "D" Then           ' Entitlements entered in days
        dblNewEntitle# = dblNewEntitle# * dblDHours#
    End If
    If rsEntRules("EH_TYPE") = "H" Then           ' Entitlements entered in Hours
        dblNewEntitle# = dblNewEntitle#
    End If
    If rsEntRules("EH_TYPE") = "F" Then           ' Entitlements entered in FTE
        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
    End If

    SQLQ = "SELECT EN_EMPNBR,EN_TYPE,EN_ID ,"
    SQLQ = SQLQ & " EN_ENTITLE, EN_TDATE FROM HR_NEXTENT_WRK "
    SQLQ = SQLQ & " WHERE EN_EMPNBR = " & snapHourEntitle("ED_EMPNBR")
    SQLQ = SQLQ & " AND EN_TYPE = '" & rsEntRules("EH_TYPE") & "'"
    SQLQ = SQLQ & " AND EN_TDATE = " & Date_SQL(DateAdd("yyyy", 1, rsEntRules("EH_TDATE")))


    snapDuplic.Open SQLQ, gdbAdoIhr001W, adOpenKeyset
    If Not snapDuplic.EOF And Not snapDuplic.BOF Then
        snapDuplic.MoveLast
    End If

    NumRec = snapDuplic.RecordCount
    If snapDuplic.EOF Then
        oldEntitleUpd = 0
    Else
        oldEntitleUpd = snapDuplic("EN_ENTITLE")
    End If
    If Accum = True Then
      If NumRec > 0 Then
        dblEntitleUpd = snapDuplic("EN_ENTITLE")
      Else
        dblEntitleUpd = 0
      End If
    Else
      dblEntitleUpd = 0
    End If

    snapDuplic.Close
    If Accum = True Then
        dblEntitleUpd = dblEntitleUpd + dblNewEntitle
    Else
        dblEntitleUpd = dblNewEntitle
    End If

    DtTm = Now


If Accum = True Then
    If NumRec > 0 Then  'if accumulate and found duplicate record

        SQLQ = "UPDATE HR_NEXTENT_WRK "
        SQLQ = SQLQ & " SET EN_ENTITLE = " & dblEntitleUpd & " "
        SQLQ = SQLQ & " WHERE EN_EMPNBR = " & snapHourEntitle("ED_EMPNBR")
        SQLQ = SQLQ & " AND EN_TYPE = '" & rsEntRules("EH_HETYPE") & "' "
        SQLQ = SQLQ & " AND EN_TDATE = " & Date_SQL(DateAdd("yyyy", 1, rsEntRules("EH_TDATE")))

        gdbAdoIhr001W.Execute (SQLQ)
    Else
        snapFuture.AddNew
        snapFuture("EN_COMPNO") = "001"
        snapFuture("EN_EMPNBR") = EmpNo&
        snapFuture("EN_ENTSORT") = 3
        snapFuture("EN_TYPE_TABL") = "ENTT"
        snapFuture("EN_TYPE") = rsEntRules("EH_HETYPE")
        snapFuture("EN_DESC") = rsEntRules("TB_DESC")
        snapFuture("EN_FDATE") = dlpDateRange(0).Text
        snapFuture("EN_TDATE") = dlpDateRange(1).Text
        snapFuture("EN_WRKEMP") = glbUserID
        snapFuture("EN_ENTITLE") = dblEntitleUpd
        snapFuture("EN_DHRS") = snapHourEntitle("ED_DHRS")
        snapFuture("EN_SURNAME") = snapHourEntitle("ED_SURNAME")
        snapFuture("EN_FNAME") = snapHourEntitle("ED_FNAME")
        snapFuture("EN_LDATE") = Now
        snapFuture("EN_LTIME") = Time$
        snapFuture("EN_LUSER") = glbUserID
        snapFuture.Update
    End If
Else
    SQLQ$ = "DELETE FROM HR_NEXTENT_WRK "
    SQLQ = SQLQ & " WHERE EN_EMPNBR = " & snapHourEntitle("ED_EMPNBR")
    SQLQ = SQLQ & " AND EN_TYPE = '" & rsEntRules("EH_HETYPE") & "' "
    SQLQ = SQLQ & " AND EN_TDATE = " & Date_SQL(DateAdd("yyyy", 1, rsEntRules("EH_TDATE")))
    
    gdbAdoIhr001W.Execute SQLQ
    
    snapFuture.AddNew
    snapFuture("EN_COMPNO") = "001"
    snapFuture("EN_EMPNBR") = EmpNo&
    snapFuture("EN_ENTSORT") = 3
    snapFuture("EN_TYPE_TABL") = "ENTT"
    snapFuture("EN_TYPE") = rsEntRules("EH_HETYPE")
    snapFuture("EN_DESC") = rsEntRules("TB_DESC")
    snapFuture("EN_FDATE") = dlpDateRange(0).Text
    snapFuture("EN_TDATE") = dlpDateRange(1).Text
    snapFuture("EN_WRKEMP") = glbUserID
    snapFuture("EN_ENTITLE") = dblEntitleUpd
    snapFuture("EN_DHRS") = snapHourEntitle("ED_DHRS")
    snapFuture("EN_SURNAME") = snapHourEntitle("ED_SURNAME")
    snapFuture("EN_FNAME") = snapHourEntitle("ED_FNAME")
    snapFuture("EN_LDATE") = Now
    snapFuture("EN_LTIME") = Time$
    snapFuture("EN_LUSER") = glbUserID
    snapFuture.Update
    snapFuture.Update
End If

DoEvents

lblNext2Rec:

MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
CommitTrans

snapFuture.Close

Screen.MousePointer = DEFAULT

Exit Sub

modannhours_Err:

If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
   'MsgBox "Conflicting Dates"
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "InsertEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub subNextHourly()
Dim c As Long, SQLQ As String

On Error GoTo Mod_Err

c = 1
'Get Entitlements that apply to today.
SQLQ = "SELECT HR_HOURLYENT.*, HRTABL.TB_DESC AS TB_DESC "
SQLQ = SQLQ & "FROM HR_HOURLYENT LEFT OUTER JOIN HRTABL ON HR_HOURLYENT.EH_HETYPE_TABL = HRTABL.TB_NAME "
SQLQ = SQLQ & "AND HR_HOURLYENT.EH_HETYPE = HRTABL.TB_KEY WHERE EH_FDATE <=" & Date_SQL(Now) & " AND EH_TDATE >=" & Date_SQL(Now) & " ORDER BY EH_ID"
rsEntRules.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
If rsEntRules.EOF = False And rsEntRules.BOF = False Then
    rsEntRules.MoveFirst
    Do
        If Not CR_SnapHour() Then Exit Sub ' create snapEntitle (form level recordset) for existing entitlments
        If snapHourEntitle.EOF = False And snapHourEntitle.BOF = False Then
            While Not snapHourEntitle.EOF
                Call modAnnHours
                snapHourEntitle.MoveNext
            Wend
        End If
        rsEntRules.MoveNext
    Loop Until rsEntRules.EOF
    snapHourEntitle.Close
End If

rsEntRules.Close
Screen.MousePointer = DEFAULT

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "subNextHourly", "Hourly", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub subNextSick()
On Error GoTo Eh
    Dim SQLQ As String
    Dim xRunTimes As Long, blIsLast As Boolean, lngRecs As Long
    
    SQLQ = "SELECT * FROM HRSICKENT ORDER BY VE_ID"
    rsEntRules.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsEntRules.EOF = False And rsEntRules.BOF = False Then
        rsEntRules.MoveFirst
        Do
                
            If Not CR_SnapEntitle("SICK") Then Exit Sub ' create snapEntitle (form level recordset) for existing entitlments
            
            If (UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N") And Not (glbCompSerial = "S/N - 2355W" And rsEntRules("VE_MANUAL") <> 0) Then
                 
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        If IsNull(snapEntitle("ED_EFDATES")) = False Then
                            'If Not IsNull(rsEntRules("VE_EDATE")) Then  'Ticket #14083
                            '    'If not a valid date then use From Date entered on this report screen
                            '    If IsDate(rsEntRules("VE_EDATE")) Then
                            '        fglbAsOf = rsEntRules("VE_EDATE")
                            '    Else
                            '        fglbAsOf = dlpDateRange(0).Text
                            '    End If
                            'Else
                            '    fglbAsOf = dlpDateRange(0).Text
                                fglbAsOf = dlpEffDate.Text
                            'End If
                            If chkAnnualized.Value = False Then
                                'fglbToDate = CDate(month(fglbAsOf) & "/" & getEOM(month(fglbAsOf)) & "/" & Year(fglbAsOf))
                                fglbToDate = DateAdd("m", 1, dlpDateRange(0).Text) - 1
                            Else
                                fglbToDate = dlpDateRange(1).Text
                            End If
                        Else
                            GoTo Fvac1
                        End If
                        For xRunTimes = 1 To 12
                            blIsLast = False
                            If xRunTimes = 12 Then blIsLast = True
                            If Not modAnnSick() Then GoTo Fvac1

                            fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
                            If chkAnnualized.Value = False Then
                                'fglbToDate = CDate(month(fglbAsOf) & "/" & getEOM(month(fglbAsOf)) & "/" & Year(fglbAsOf))
                                fglbToDate = CDate(month(DateAdd("m", 1, fglbToDate)) & "/" & getEOM(month(DateAdd("m", 1, fglbToDate))) & "/" & Year(DateAdd("m", 1, fglbToDate)))
                            Else
                                fglbToDate = dlpDateRange(1).Text
                            End If
                            
                            If fglbAsOf > dlpDateRange(1).Text Then Exit For
                        Next
                        
Fvac1:
                        snapEntitle.MoveNext
                    Wend
                End If
                snapEntitle.Close
            Else
                'create snapEntitle (form level recordset)
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount

                        If IsNull(snapEntitle("ED_EFDATES")) = False Then
                            'If Not IsNull(rsEntRules("VE_EDATE")) Then  'Ticket #14083
                            '    'If not a valid date then use From Date entered on this report screen
                            '    If IsDate(rsEntRules("VE_EDATE")) Then
                            '        fglbAsOf = rsEntRules("VE_EDATE")
                            '    Else
                            '        fglbAsOf = dlpDateRange(0).Text
                            '    End If
                            'Else
                            '    fglbAsOf = dlpDateRange(0).Text
                                fglbAsOf = dlpEffDate.Text
                            'End If
                            fglbToDate = dlpDateRange(1).Text
                        Else
                            GoTo Fvac2
                        End If

                        If Not modAnnSick() Then GoTo Fvac2
                        
Fvac2:
                        snapEntitle.MoveNext
                    Wend

                End If
                snapEntitle.Close
            End If
            rsEntRules.MoveNext

        Loop Until rsEntRules.EOF
    End If

    rsEntRules.Close


exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "subNextVac", "WORK File", "CREATE")
    If gintRollBack% = False Then Resume exH Else Unload Me

End Sub

Private Function modAnnSick() As Boolean
Dim EmpNo As Long
Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, fglbWDate As String
Dim xAsOf

Dim xComments
On Error GoTo modUpdateSelection_Err
modAnnSick = False

gdbAdoIhr001.BeginTrans

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select

    EmpNo& = snapEntitle("ED_EMPNBR")
    
    'if no records exist for this employee create one
    If CR_snapFuture(EmpNo, "SIC") = 0 Or _
        ((UCase(glbCompEntSick$) = "M" Or UCase(glbCompEntSick$) = "N") And Not _
        (glbCompSerial = "S/N - 2355W" And rsEntRules("VE_MANUAL") <> 0) And chkAnnualized.Value = False) Then
        
        snapFuture.AddNew
        snapFuture("EN_COMPNO") = "001"
        snapFuture("EN_EMPNBR") = EmpNo&
        snapFuture("EN_ENTSORT") = 2
        snapFuture("EN_TYPE_TABL") = "ENTT"
        snapFuture("EN_TYPE") = "SIC"
        snapFuture("EN_DESC") = "Sick"
        
        If Day(DateAdd("m", -1, fglbToDate) + 1) = 31 Then
            snapFuture("EN_FDATE") = DateAdd("m", -1, fglbToDate) + 2
        ElseIf Day(DateAdd("m", -1, fglbToDate) + 1) = 29 And month(fglbToDate) = 2 Then
            snapFuture("EN_FDATE") = CVDate(month(fglbToDate) & "/1/" & Year(fglbToDate))
        Else
            snapFuture("EN_FDATE") = DateAdd("m", -1, fglbToDate) + 1
        End If
        'snapFuture("EN_FDATE") = fglbAsOf
        
        snapFuture("EN_TDATE") = fglbToDate
        snapFuture("EN_SURNAME") = snapEntitle("ED_SURNAME")
        snapFuture("EN_FNAME") = snapEntitle("ED_FNAME")
        snapFuture("EN_WRKEMP") = glbUserID
        snapFuture("EN_LDATE") = Now
        snapFuture("EN_LTIME") = Time$
        snapFuture("EN_LUSER") = glbUserID
        snapFuture.Update
    End If
    
    If IsNull(snapFuture("EN_ENTITLE")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapFuture("EN_ENTITLE")
    End If
    
  
    'Ticket #22206 - Not to include Previous in the future entitlement calculation.
    'If IsNull(snapEntitle("ED_PSICK")) Then
        dblPrevEntitle# = 0
    'Else
    '    dblPrevEntitle# = snapEntitle("ED_PSICK")
    'End If
    
    'Ticket #22206 - Not to include Taken in the future entitlement calculation.
    'If IsNull(snapEntitle("ED_SICKT")) Then
        dblTKEEntitle# = 0
    'Else
    '    dblTKEEntitle# = snapEntitle("ED_SICKT")
    'End If
    
    spt = snapEntitle("ED_PT")
    
    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)
    
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    rsJOB.Close
    If glbLinamar Then dblDHours# = 8
    
    xAsOf = fglbAsOf
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    intWhereFit& = -1


    If rsEntRules("VE_EMONTH") > 0 Then
        If dblServiceYears# >= CDbl(rsEntRules("VE_BMONTH")) And dblServiceYears# <= CDbl(rsEntRules("VE_EMONTH")) Then
            intWhereFit& = x%
        End If
    End If

    
    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    

    dblNewEntitle# = rsEntRules("VE_ENTITLE")
    dblNewMax# = 0
    If rsEntRules("VE_TYPE") = "D" Then           ' Entitlements entered in days
        If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblDHours#
        dblNewEntitle# = dblNewEntitle# * dblDHours#
        dblEntitleUpd = dblNewEntitle
    End If
    If rsEntRules("VE_TYPE") = "F" Then
        If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblFTEHours# * dblDHours#
        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
        dblEntitleUpd = dblNewEntitle
    End If
    If rsEntRules("VE_TYPE") = "H" Then
        If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX")
    End If

    dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values

    
    If dblNewMax <> 0 Then          'only do if not zero
        'Ticket #22206 - Not to include Taken in the future entitlement calculation.
        'If glbCompSerial = "S/N - 2173W" Or glbCompSerial = "S/N - 2375W" Then 'for town of Ajax or City of Timmins
        '    If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
        '        dblEntitleUpd = dblEntitle#
        '    ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
        '        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
        '    End If
        'Else
            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                dblEntitleUpd = dblNewMax - dblPrevEntitle#
            End If
        'End If
    End If
    
    If glbCBrant Then
        If snapEntitle("ED_HIRECODE") = "Y" And dblTKEEntitle# > 0 Then
            dblEntitleUpd = dblEntitleUpd - dblTKEEntitle#
        End If
    End If
    DtTm = Now
    
   snapFuture("EN_ENTITLE") = dblEntitleUpd
   snapFuture("EN_DHRS") = dblDHours#
   snapFuture.Update
lblNextRec:
    gdbAdoIhr001.CommitTrans
    DoEvents

snapFuture.Close

modAnnSick = True


Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modAnnSick", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Function XLWriter()
    Dim SQLQ As String
    'Dim exApp As Excel.Application, exBook As Excel.Workbook, exSheet As Excel.Worksheet
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim xlsFileTmp As String, xlsFileMat As String
    Dim rsEnts As New ADODB.Recordset, rsEmps As New ADODB.Recordset, rsIn As New ADODB.Recordset, rsGroup As New ADODB.Recordset
    Dim xRow As Long, xCol As Long, xwCol As Long, xMax As Long
    Dim xType As String, strTemp As String, strName As String, strEmp As String
    Dim retVal As Long, strTime As String
    Dim xStartDate As String, dblSick As Double, dblVac As Double
    Dim xExcelRptPath  As String
    
    On Error GoTo Err_XLS
    
'WriteFile ("Starting...")

'WriteFile ("Acquiring template file and creating the Future Entitlement file to update with entitlements")

    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
    
    xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "FutureEntstmp.xls"
    'Ticket# 8293
    If glbLinamar Then
        xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "Reports\FutureEntsMat" & Trim(glbUserID) & ".xls"
    Else
        'Ticket #22034 - May need to save report in different path
        'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "FutureEntsMat" & Trim(glbUserID) & ".xls"
        xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "FutureEntsMat" & Trim(glbUserID) & ".xls"
    End If

    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).Caption = "Please wait..."
    
    Screen.MousePointer = HOURGLASS

    If Dir(xlsFileTmp) = "" Then
        MsgBox "There is no " & xlsFileTmp
        Exit Function
    End If
    If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat

    FileCopy xlsFileTmp, xlsFileMat
    
'WriteFile ("New Future Entitlement file ready to be updated")
    
    'Create new WorkBook of Excel
    'Set exApp = New Excel.Application      'Ticket #30219 - Was giving errors to some so switched to 'CreateObject' instead of 'New'
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(xlsFileMat)
    
'WriteFile ("Created the Workbook")

    'Create Grouping recordset
    SQLQ = getGroupField
    
'WriteFile ("Got the query statement: " & SQLQ)

    rsGroup.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    
'WriteFile ("Opened the recordset")

    If rsGroup.EOF = False And rsGroup.BOF = False Then
    
'WriteFile ("Creating Grouping: " & SQLQ)

        strTime = Time$
        Do
            If comGroup(0).Text = "(none)" Then
                Set exSheet = exBook.Sheets(1)
                exSheet.name = "All"
            Else
                exBook.Worksheets(1).Copy After:=exBook.Worksheets(exBook.Worksheets.count)
                Set exSheet = exBook.Worksheets(exBook.Worksheets.count)
                
                If IsNull(rsGroup("TB_DESC")) Then
                    exSheet.name = "None"
                Else
                    exSheet.name = Left(rsGroup("TB_DESC"), 30)
                End If
            End If
            
            'Get the codes for columns
            SQLQ = "SELECT DISTINCT HR_NEXTENT_WRK.EN_DESC,HR_NEXTENT_WRK.EN_ENTSORT,HR_NEXTENT_WRK.EN_TYPE FROM HR_NEXTENT_WRK "
            SQLQ = SQLQ & "ORDER BY HR_NEXTENT_WRK.EN_ENTSORT, HR_NEXTENT_WRK.EN_DESC "
            rsEnts.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
            xMax = rsEnts.RecordCount
            If xMax > 250 Then
                retVal = MsgBox("This report exceeds 250 entitlements, Click Yes to Continue, No to refine the query", vbYesNo + vbQuestion, "Columns Exceeded")
                If retVal = vbNo Then
                    GoTo exH
                End If
            End If
            
            'Get the Employees for the Report
            SQLQ = "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
            gdbAdoIhr001W.Execute SQLQ
            If Not glbSQL And Not glbOracle Then Call Pause(3)
            
            SQLQ = "INSERT INTO HR_EMPLIST_WRK (TT_EMPNBR, TT_WRKEMP) " & in_SQL(glbIHRDBW)
            SQLQ = SQLQ & " SELECT HREMP.ED_EMPNBR, '" & glbUserID & "' as TT_WRKEMP FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
            SQLQ = SQLQ & " WHERE (HR_JOB_HISTORY.JH_CURRENT <> 0) And " & glbstrSelCri
            If IsNull(rsGroup("grpField")) Then
                If glbSQL Or glbOracle Then
                    SQLQ = SQLQ & " and " & strGroup & " IS NULL "
                Else
                    SQLQ = SQLQ & " and isnull(" & strGroup & ") = true "
                End If
            Else
                If strGroup = "HR_JOB_HISTORY.JH_REPTAU" Then 'Supervisor is a number
                    SQLQ = SQLQ & " and " & strGroup & "=" & rsGroup("grpField") & " "
                Else
                    SQLQ = SQLQ & " and " & strGroup & "='" & Replace(rsGroup("grpField"), "'", "''") & "' "
                End If
            End If

            gdbAdoIhr001.Execute SQLQ, , adCmdText

'WriteFile ("Employees for the Report: " & SQLQ)

            If Not glbSQL And Not glbOracle Then Call Pause(3)
            
            SQLQ = "SELECT DISTINCT HR_NEXTENT_WRK.EN_EMPNBR, HR_NEXTENT_WRK.EN_SURNAME, HR_NEXTENT_WRK.EN_FNAME, HR_NEXTENT_WRK.EN_EMPNBR, "
            SQLQ = SQLQ & "HR_NEXTENT_WRK.EN_FDATE, HR_NEXTENT_WRK.EN_TDATE "
'            If chkAnnualized.Visible And chkAnnualized.Value = False Then
'                SQLQ = SQLQ & ",HR_NEXTENT_WRK.EN_ENTSORT "
'            End If
            SQLQ = SQLQ & "FROM HR_NEXTENT_WRK "
            SQLQ = SQLQ & " WHERE HR_NEXTENT_WRK.EN_WRKEMP='" & glbUserID & "' AND HR_NEXTENT_WRK.EN_EMPNBR IN "
            SQLQ = SQLQ & " (SELECT HR_EMPLIST_WRK.TT_EMPNBR FROM HR_EMPLIST_WRK WHERE HR_EMPLIST_WRK.TT_WRKEMP='" & glbUserID & "') "
            SQLQ = SQLQ & " ORDER BY "
            If comGroup(1).ListIndex = 0 Then
                SQLQ = SQLQ & "HR_NEXTENT_WRK.EN_SURNAME,HR_NEXTENT_WRK.EN_FNAME, "
            End If
'            If chkAnnualized.Visible And chkAnnualized.Value = False Then
'                SQLQ = SQLQ & "HR_NEXTENT_WRK.EN_ENTSORT, "
'            End If
            SQLQ = SQLQ & " HR_NEXTENT_WRK.EN_FDATE, HR_NEXTENT_WRK.EN_TDATE "
            rsEmps.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
       
'WriteFile ("Updating Future Entitlement file")
            
            'Sheet Labels
            exSheet.Cells(2, 1) = "Date: " & Format(Now, "mmm dd, yyyy")
            exSheet.Cells(3, 1) = "Time: " & strTime
            exSheet.Cells(5, 1) = "Group by " & comGroup(0).Text
            
            If rsEmps.EOF = False And rsEmps.BOF = False Then
        '*********************************
                
                xCol = 4
                Do
                    xRow = 10
                    
                    exSheet.Cells(9, xCol) = rsEnts("EN_DESC")
                    'exSheet.Range(exSheet.Cells(8, xCol), exSheet.Cells(8, xCol + 1)).Merge
                    'exSheet.Cells(9, xCol) = "Entitlement"
                    'exSheet.Cells(9, xCol + 1) = "Taken"
                    rsEmps.MoveFirst
                    strTemp = ""
                    strName = ""
                    strEmp = rsEmps("EN_SURNAME") & ", " & rsEmps("EN_FNAME")
                    dblVac = 0
                    dblSick = 0
                    Do
                        If xCol = 4 Then
                            If strName <> (rsEmps("EN_SURNAME") & ", " & rsEmps("EN_FNAME")) Then
                                strName = rsEmps("EN_SURNAME") & ", " & rsEmps("EN_FNAME")
                                exSheet.Cells(xRow, 1) = strName
                            End If
                            
                            If strTemp <> (rsEmps("EN_FDATE") & " - " & rsEmps("EN_TDATE")) Then
                                strTemp = Format(rsEmps("EN_FDATE"), "mm/dd/yyyy") & " - " & Format(rsEmps("EN_TDATE"), "mm/dd/yyyy")
                                exSheet.Cells(xRow, 2) = strTemp
                            End If
                        End If
                        
                        SQLQ = "SELECT EN_ENTITLE, EN_TAKEN, EN_TYPE, EN_DHRS FROM HR_NEXTENT_WRK WHERE EN_EMPNBR=" & rsEmps("EN_EMPNBR")
                        SQLQ = SQLQ & " AND EN_FDATE=" & Date_SQL(rsEmps("EN_FDATE")) & " AND EN_TDATE=" & Date_SQL(rsEmps("EN_TDATE"))
                        SQLQ = SQLQ & " AND EN_WRKEMP='" & glbUserID & "'"  'ORDER BY  HR_NEXTENT_WRK.EN_FDATE, HR_NEXTENT_WRK.EN_TDATE"
                        SQLQ = SQLQ & "  AND EN_ENTITLE IS NOT NULL ORDER BY HR_NEXTENT_WRK.EN_FDATE, HR_NEXTENT_WRK.EN_TDATE"
                        rsIn.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic, adCmdText
                        If rsIn.EOF = False And rsIn.BOF = False Then
                            Do
                                If rsEnts("EN_TYPE") = rsIn("EN_TYPE") Then
                                    If cmbHours.ListIndex = 0 Then
                                        exSheet.Cells(xRow, xCol) = rsIn("EN_ENTITLE")
                                        If Not IsNull(rsIn("EN_ENTITLE")) Then
                                            If rsEnts("EN_TYPE") = "VAC" Then
                                                dblVac = dblVac + Val(rsIn("EN_ENTITLE"))
                                            ElseIf rsEnts("EN_TYPE") = "SIC" Then
                                                dblSick = dblSick + Val(rsIn("EN_ENTITLE"))
                                            End If
                                        End If
                                    Else
                                        If rsIn("EN_DHRS") = 0 Then
                                            exSheet.Cells(xRow, xCol) = "0.00"
                                            If rsEnts("EN_TYPE") = "VAC" Then
                                                dblVac = dblVac + 0
                                            ElseIf rsEnts("EN_TYPE") = "SIC" Then
                                                dblSick = dblSick + 0
                                            End If
                                        Else
                                            If Not IsNull(rsIn("EN_ENTITLE")) Then
                                                exSheet.Cells(xRow, xCol) = rsIn("EN_ENTITLE") / rsIn("EN_DHRS")
                                                If rsEnts("EN_TYPE") = "VAC" Then
                                                    dblVac = dblVac + (rsIn("EN_ENTITLE") / rsIn("EN_DHRS"))
                                                ElseIf rsEnts("EN_TYPE") = "SIC" Then
                                                    dblSick = dblSick + (rsIn("EN_ENTITLE") / rsIn("EN_DHRS"))
                                                End If
                                            End If
                                        End If
                                        
                                    End If
                                    'exSheet.Cells(xRow, xCol + 1) = rsIn("EN_TAKEN")
                                End If
                                rsIn.MoveNext
                            Loop Until rsIn.EOF
                        End If
                        rsIn.Close
                        xRow = xRow + 1
                        rsEmps.MoveNext
                        'Total
                        If (chkAnnualized.Visible = True And chkAnnualized.Value = False) Or (UCase(glbCompEntVac$) = "N" And chkVacation.Value = True) Or (UCase(glbCompEntSick$) = "N" And chkSick.Value = True) Then
                            If rsEmps.EOF Then
                                exSheet.Cells(xRow, 3) = "Total"
                                exSheet.Cells(xRow, 3).Font.Bold = True
                                If rsEnts("EN_TYPE") = "VAC" Then
                                    exSheet.Cells(xRow, xCol) = dblVac
                                ElseIf rsEnts("EN_TYPE") = "SIC" Then
                                    exSheet.Cells(xRow, xCol) = dblSick
                                End If
                                exSheet.Cells(xRow, xCol).Font.Bold = True
                                dblVac = 0
                                dblSick = 0
                            ElseIf strEmp <> rsEmps("EN_SURNAME") & ", " & rsEmps("EN_FNAME") Then
                                exSheet.Cells(xRow, 3) = "Total"
                                exSheet.Cells(xRow, 3).Font.Bold = True
                                If rsEnts("EN_TYPE") = "VAC" Then
                                    exSheet.Cells(xRow, xCol) = dblVac
                                ElseIf rsEnts("EN_TYPE") = "SIC" Then
                                    exSheet.Cells(xRow, xCol) = dblSick
                                End If
                                exSheet.Cells(xRow, xCol).Font.Bold = True
                                dblVac = 0
                                dblSick = 0
                                strEmp = rsEmps("EN_SURNAME") & ", " & rsEmps("EN_FNAME")
                                xRow = xRow + 2
                            End If
                        End If
                    Loop Until rsEmps.EOF
                    xCol = xCol + 1
                    rsEnts.MoveNext
                Loop Until rsEnts.EOF
        
'WriteFile ("Update completed. Now formating")

        '*********************************
                If chkEntitle.Value = False Then
                    exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(1, 6)).Merge
                    exSheet.PageSetup.PrintArea = exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(xRow, 6)).AddressLocal
                ElseIf xCol < 7 Then
                    exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(1, 7)).Merge
                    exSheet.PageSetup.PrintArea = exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(xRow, 7)).AddressLocal
                Else
                    exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(1, xCol)).Merge
                    exSheet.PageSetup.PrintArea = exSheet.Range(exSheet.Cells(1, 1), exSheet.Cells(xRow, xCol)).AddressLocal
                End If
                
                exSheet.Range(exSheet.Cells(10, 4), exSheet.Cells(xRow, xCol)).NumberFormat = "#0.00"
                
                
                exSheet.PageSetup.Orientation = xlLandscape
                exSheet.PageSetup.Zoom = False
                exSheet.Range("A1").Select
                exBook.Save
                Set exSheet = Nothing
                                   
'WriteFile ("Formating complete. Saved the Workbook")

            End If
            rsEmps.Close
            rsEnts.Close
            
            rsGroup.MoveNext
        Loop Until rsGroup.EOF
        rsGroup.Close
        Set rsGroup = Nothing
        
'WriteFile ("Grouping Loop complete")

        Set exSheet = exBook.Worksheets(1)
        
'WriteFile ("Display Alert = False")

        exApp.DisplayAlerts = False
        
'WriteFile ("Delete Worksheet 1")

        exSheet.Delete
        
'WriteFile ("Display Alert = True")

        exApp.DisplayAlerts = True
        
'WriteFile ("Set Worksheet 2 now as Active Worksheet 1")

        exBook.Worksheets(1).Activate
        
'WriteFile ("Save the Workbook")

        exBook.Save
        
'WriteFile ("Set Workbook to nothing")

        Set exBook = Nothing
        
'WriteFile ("Set Application visible to True")

        exApp.Visible = True
        
'WriteFile ("Set Appliation to Nothing")

        Set exApp = Nothing
        
    Else
        MsgBox "There are no Records matching this criteria", vbInformation + vbOKOnly, "No Records"
        exBook.Close False
        exApp.Quit
        Set exApp = Nothing
    End If
    
'WriteFile ("Done")

    Call Pause(1)
    
exH:
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    Screen.MousePointer = DEFAULT
    Exit Function
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
        Exit Function
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
        Exit Function
    End If
    If Err = 76 Then
        MsgBox Err.Description & " to save the Training Matrix Report." & vbCrLf & "Please check Company Preference under Setup Menu."
        Exit Function
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
      
End Function

Private Sub chkEntitle_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbHours_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
    x% = Cri_SetAll()
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    MDIMain.Timer1.Enabled = True
     Call set_PrintState(True)

End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox CStr(Err.Number) + ": " + Err.Description, vbExclamation + vbOKOnly, "Err cmdView"

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
    Screen.MousePointer = HOURGLASS
     Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
     
    x% = Cri_SetAll()
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)

End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox CStr(Err.Number) + ": " + Err.Description, vbExclamation + vbOKOnly, "Err cmdView"

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
    Case 1: strCd$ = "HREMP.ED_ORG"
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    Case 6: strCd$ = "HREMP.ED_SALDIST"

    End Select
        CodeCri = "(" & strCd$ & " in  ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "'))"
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
Dim countr   As Integer

If Len(clpDiv.Text) > 0 Then
    DivCri = "(HREMP.ED_DIV in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
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

Private Sub Cri_Dept()
Dim DivCri As String
Dim countr   As Integer

If Len(clpDept.Text) > 0 Then
    DivCri = "(HREMP.ED_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "'))"
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
    EECri = "HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
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

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpReptau.Text) > 0 Then
    EECri = "HR_JOB_HISTORY.JH_REPTAU IN (" & getEmpnbr(elpReptau.Text) & ") "
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

EECri = "HREMP.ED_PT in ('" & Replace(clpPT.Text, ",", "','") & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Sub Cri_FTDates()

'Dim TempCri As String
'Dim dtYYY%, dtMM%, dtDD%, X%
'Dim FromDate, ToDate, SQLQ
'Dim RsHRPARCO As New ADODB.Recordset
'
'If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
'    TempCri = "(((HREMP.ED_EFDATE >= " & Date_SQL(dlpDateRange(0).Text) & ") and "
'    TempCri = TempCri & " (HREMP.ED_ETDATE <= " & Date_SQL(dlpDateRange(1).Text) & ")) "
'End If
'
'If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
'    TempCri = TempCri & "or ((HREMP.ED_EFDATES >= " & Date_SQL(dlpDateRange(0).Text) & ") and "
'    TempCri = TempCri & " (HREMP.ED_ETDATES <= " & Date_SQL(dlpDateRange(1).Text) & "))) "
'End If
'
'
'Cri_FTDatst:
'If Len(TempCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = TempCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & TempCri
'    End If
'    glbiOneWhere = True
'End If

End Sub

Private Function Cri_SetAll()
Dim x%, xNoFile, xNoWork, SQLQ As String

Dim cmpDates As Boolean
Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = "(1=1)"

cmpDates = False

If chkSick.Value = False And chkEntitle.Value = False And chkVacation.Value = False Then
    MsgBox "No Entitlements Selected", vbInformation + vbOKOnly, "Select Entitlements"
    Exit Function
End If

Call Cri_Dept
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_Sup
For x% = 0 To 6
    Call Cri_Code(x%)
Next x%
Call Cri_FTDates

'Clear the work Table
SQLQ = "DELETE From HR_NEXTENT_WRK "
SQLQ = SQLQ & "WHERE EN_WRKEMP = '" & glbUserID & "'"
gdbAdoIhr001W.Execute SQLQ

'Find Next Vacation Entitlement
If chkVacation.Value Then Call subNextVac

'Find Next Sick Entitlement
If chkSick.Value Then Call subNextSick

'Find Hourly Entitlements
If chkEntitle.Value = True Then Call subNextHourly

Cri_SetAll = XLWriter

Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Entitlement Report", "Cri_SetAll", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, x%

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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, x%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "HREMP.ED_SHIFT= '" & txtShift.Text & "'"

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
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0
Cri_Sorts = z% ' next section number to format

Me.vbxCrystal.Formulas(1) = "lblTitle='Future Entitlements'"

End Function

Private Function CriCheck()
Dim x%

CriCheck = False

If IsDate(dlpDateRange(0).Text) = False Or IsDate(dlpDateRange(1).Text) = False Then
    MsgBox "Future Entitlement Range must be a entered", , "info:HR"
    dlpDateRange(0).SetFocus
    Exit Function
End If

If DateDiff("d", dlpDateRange(0).Text, dlpDateRange(1).Text) < 0 Then
    MsgBox "Future Entitlement To date must be after From date", , "info:HR"
    dlpDateRange(0).SetFocus
    Exit Function
End If

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known"), , "info:HR"
    'clpDiv.SetFocus
    Exit Function
End If

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known", , "info:HR"
    'clpDept.SetFocus
    Exit Function
End If

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Len(dlpEffDate.Text) > 0 Then
    If Not IsDate(dlpEffDate.Text) Then
        MsgBox "Invalid Effective Date"
        dlpEffDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "Effective Date is required"
    dlpEffDate.SetFocus
    Exit Function
End If

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

glbOnTop = "FRMRNEXTENT"

Screen.MousePointer = HOURGLASS

Me.WindowState = vbMaximized

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call setRptCaption(Me)
Call setCaption(lblSalDist)
'Call setCaption(lblRepAuth)    'Ticket #27539 - It's Reporting Authority 1 from Current Position and not Attendance Supervisor
Call comGrpLoad

'Ticket #27539 - It's Reporting Authority 1 from Current Position and not Attendance Supervisor
'If lblRepAuth.Caption = "AttSupervisor" Then lblRepAuth.Caption = "Supervisor"

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbLinamar Then clpCode(3).MaxLength = 8

If glbCompSerial = "S/N - 2235W" Then   'laura 03/09/98
    lblTitle(5).Visible = False
    cmbHours.Visible = False
ElseIf glbCompSerial = "S/N - 2236W" Then
    lblTitle(5).Visible = False
    cmbHours.Visible = False
Else
    lblTitle(5).Visible = True
    cmbHours.Visible = True
End If

If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    chkEntitle.Visible = False
    chkSick.Visible = False
End If
If glbLinamar Then
    lblSalDist.Visible = True
    clpCode(6).Visible = True
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

lblFromTo.Visible = True
dlpDateRange(0).Visible = True
dlpDateRange(1).Visible = True

'If the company master is Monthly entitlement give the option of only showing the total
chkAnnualized.Visible = getcMaster()

chkEntitle = False

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 8910 Then
        scrControl.Value = 0
        scrFrame.Top = 120
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 4000 Then
            scrControl.Max = 5000
        Else
            scrControl.Max = 3500
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If

    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 9615 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 150
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 120
    End If
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmRNextEnt = Nothing
End Sub

Private Sub comGrpLoad()

comGroup(0).AddItem lStr("Division")
comGroup(0).AddItem lStr("Department")
comGroup(0).AddItem lStr("Location")
comGroup(0).AddItem lStr("Union")
comGroup(0).AddItem lStr("Administered By")
comGroup(0).AddItem "Employee Name"
comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000

If glbLinamar Then ' Frank May 2,2001
    comGroup(0).AddItem "Employment Type"
    comGroup(0).AddItem ("Home Line")
End If

If Not glbMulti Then comGroup(0).AddItem "Shift"

comGroup(0).AddItem lStr("Rept. Authority 1")
comGroup(0).AddItem lStr("Region")
comGroup(0).AddItem "(none)"
comGroup(0).ListIndex = 0
comGroup(1).AddItem "Employee Name"
comGroup(1).ListIndex = 0
comGroup(1).Enabled = False

cmbHours.AddItem "Hours"
cmbHours.AddItem "Days"
cmbHours.ListIndex = 0

End Sub

Private Sub scrControl_Change()
    scrFrame.Top = 120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
    scrFrame.Left = 40 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
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

Private Sub subNextVac()
    On Error GoTo Eh
    Dim SQLQ As String
    Dim xRunTimes As Long, blIsLast As Boolean, lngRecs As Long
    
    SQLQ = "SELECT * FROM HRVACENT ORDER BY VE_ID"
    rsEntRules.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rsEntRules.EOF = False And rsEntRules.BOF = False Then
        rsEntRules.MoveFirst
        Do
                
            If Not CR_SnapEntitle("VAC") Then Exit Sub ' create snapEntitle (form level recordset) for existing entitlments
            
            If (UCase(glbCompEntVac$) = "M" Or UCase(glbCompEntVac$) = "N") And Not (glbCompSerial = "S/N - 2355W" And rsEntRules("VE_MANUAL") <> 0) Then
                 
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        If IsNull(snapEntitle("ED_EFDATE")) = False Then
                            'If Not IsNull(rsEntRules("VE_EDATE")) Then  'Ticket #14083
                            '    'If not a valid date then use From Date entered on this report screen
                            '    If IsDate(rsEntRules("VE_EDATE")) Then
                            '        fglbAsOf = rsEntRules("VE_EDATE")
                            '    Else
                            '        fglbAsOf = dlpDateRange(0).Text
                            '    End If
                            'Else
                                'fglbAsOf = dlpDateRange(0).Text
                                fglbAsOf = dlpEffDate.Text
                            'End If
                            If chkAnnualized.Value = False Then
                                'fglbToDate = CDate(month(fglbAsOf) & "/" & getEOM(month(fglbAsOf)) & "/" & Year(fglbAsOf))
                                fglbToDate = DateAdd("m", 1, dlpDateRange(0).Text) - 1
                            Else
                                fglbToDate = dlpDateRange(1).Text
                            End If
                        Else
                            GoTo Fvac1
                        End If
                        For xRunTimes = 1 To 12
                            blIsLast = False
                            If xRunTimes = 12 Then blIsLast = True
                            If Not modAnnVacation(blIsLast) Then GoTo Fvac1

                            fglbAsOf = DateAdd("m", 1, CVDate(fglbAsOf))
                            If chkAnnualized.Value = False Then
                                'fglbToDate = CDate(month(fglbAsOf) & "/" & getEOM(month(fglbAsOf)) & "/" & Year(fglbAsOf))
                                fglbToDate = CDate(month(DateAdd("m", 1, fglbToDate)) & "/" & getEOM(month(DateAdd("m", 1, fglbToDate))) & "/" & Year(DateAdd("m", 1, fglbToDate)))
                            Else
                                fglbToDate = dlpDateRange(1).Text
                            End If
                            
                            If DateDiff("d", fglbAsOf, dlpDateRange(1).Text) < 0 Then Exit For
                        Next
Fvac1:
                        snapEntitle.MoveNext
                    Wend
                End If
                snapEntitle.Close
            Else
                'create snapEntitle (form level recordset)
                If snapEntitle.EOF = False And snapEntitle.BOF = False Then
                    While Not snapEntitle.EOF
                        lngRecs = snapEntitle.RecordCount

                        If IsNull(snapEntitle("ED_EFDATE")) = False Then
                            'If Not IsNull(rsEntRules("VE_EDATE")) Then  'Ticket #14083
                            '    'If not a valid date then use From Date entered on this report screen
                            '    If IsDate(rsEntRules("VE_EDATE")) Then
                            '        fglbAsOf = rsEntRules("VE_EDATE")
                            '    Else
                            '        fglbAsOf = dlpDateRange(0).Text
                            '    End If
                            'Else
                            '    fglbAsOf = dlpDateRange(0).Text
                                fglbAsOf = dlpEffDate.Text
                            'End If
                            fglbToDate = dlpDateRange(1).Text
                        Else
                            GoTo Fvac2
                        End If

                        If Not modAnnVacation(True) Then GoTo Fvac2
Fvac2:
                        snapEntitle.MoveNext
                    Wend

                End If
                snapEntitle.Close
            End If
            rsEntRules.MoveNext

        Loop Until rsEntRules.EOF
    End If

    rsEntRules.Close

exH:
    Exit Sub
Eh:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "subNextVac", "WORK File", "CREATE")
    If gintRollBack% = False Then Resume exH Else Unload Me

End Sub

Private Function CR_SnapEntitle(Optional xType)
Dim SQLQ As String

CR_SnapEntitle = False
On Error GoTo CR_SnapEntitle_Err

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_FNAME,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_ANNVAC, ED_ANNSICK, ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES,ED_SICKT,"
SQLQ = SQLQ & " ED_DIV,ED_PT, ED_SECTION,ED_SALDIST, ED_LOC, ED_ORG, ED_EMP,"
SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME "

If glbOracle Then
    SQLQ = SQLQ & "FROM HREMP, HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT<>0"
Else
    SQLQ = SQLQ & "FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0"
End If

If Not IsMissing(xType) Then
    SQLQ = SQLQ & " AND " & getWSQLQ(xType)
Else
    SQLQ = SQLQ & " AND " & getWSQLQ("")
End If

If Len(rsEntRules("VE_GRPCD")) > 0 Then
    SQLQ = SQLQ & " AND ED_EMPNBR IN "
    SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
    SQLQ = SQLQ & " WHERE JB_GRPCD = '" & rsEntRules("VE_GRPCD") & "') "
End If
SQLQ = SQLQ & " AND " & glbstrSelCri

If snapEntitle.State <> 0 Then snapEntitle.Close
snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockPessimistic

CR_SnapEntitle = True

Exit Function

CR_SnapEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function getWSQLQ(xType) As String
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSection
Dim xFromDate
Dim xToDate, fglbESQLQ As String

fglbESQLQ = glbSeleDeptUn
If Len(rsEntRules("VE_DEPT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND  ED_DEPTNO = '" & rsEntRules("VE_DEPT") & "' "
If Len(rsEntRules("VE_DIV")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & rsEntRules("VE_DIV") & "' "
If Len(rsEntRules("VE_ORG")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & rsEntRules("VE_ORG") & "' "
If Len(rsEntRules("VE_EMP")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & rsEntRules("VE_EMP") & "' "
If glbLinamar Then
    If Len(rsEntRules("VE_SECTION")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SALDIST = '" & rsEntRules("VE_SECTION") & "' "
Else
    If Not glbCBrant Then 'added by Bryan 18/Apr/2006 Ticket#10495
        If Len(rsEntRules("VE_SECTION")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & rsEntRules("VE_SECTION") & "' "
    End If
End If
If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
    If xType = "VAC" Then
        If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM1 = '" & rsEntRules("VE_LOC") & "' "
    ElseIf xType = "SICK" Then
        If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_VADIM2 = '" & rsEntRules("VE_LOC") & "' "
    End If
Else
    If Len(rsEntRules("VE_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & rsEntRules("VE_LOC") & "' "
End If

If Len(rsEntRules("VE_PT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & rsEntRules("VE_PT") & "' "

getWSQLQ = fglbESQLQ

End Function

Private Function getWSQLQ_HR(xType) As String
Dim SQLQ As String, fglbESQLQ As String

fglbESQLQ = glbSeleDeptUn
If Len(rsEntRules("EH_DEPT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DEPTNO = '" & rsEntRules("EH_DEPT") & "'"
If Len(rsEntRules("EH_DIV")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_DIV = '" & rsEntRules("EH_DIV") & "' "
If Len(rsEntRules("EH_ORG")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_ORG = '" & rsEntRules("EH_ORG") & "' "
If Len(rsEntRules("EH_EMP")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_EMP = '" & rsEntRules("EH_EMP") & "' "
If Len(rsEntRules("EH_SECTION")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_SECTION = '" & rsEntRules("EH_SECTION") & "' "
If glbCompSerial = "S/N - 2380W" Then 'VitalAire Ticket #12591
    If Len(rsEntRules("EH_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PROV = '" & rsEntRules("EH_LOC") & "' "
Else
    If Len(rsEntRules("EH_LOC")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_LOC = '" & rsEntRules("EH_LOC") & "' "
End If
If Len(rsEntRules("EH_PT")) > 0 Then fglbESQLQ = fglbESQLQ & " AND ED_PT = '" & rsEntRules("EH_PT") & "' "

getWSQLQ_HR = fglbESQLQ

End Function

Private Function modAnnVacation(isLast As Boolean)
Dim EmpNo As Long
Dim dblEntitle#, dblPrevEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim xAsOf, fglbWDate As String
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim xComments
Dim dblEntitleDays

On Error GoTo modUpdateSelection_Err

modAnnVacation = False

gdbAdoIhr001.BeginTrans

    if_Entitle = False
    if_Vacation = False

Select Case glbCompWDate$ ' sets field reference for basic 'which date'
    Case "O": fglbWDate$ = "ED_DOH"
    Case "S": fglbWDate$ = "ED_SENDTE"
    Case "U": fglbWDate$ = "ED_UNION"
    Case "L": fglbWDate$ = "ED_LTHIRE"
    Case "D": fglbWDate$ = "ED_USRDAT1"
End Select

    EmpNo = snapEntitle("ED_EMPNBR")

    'if no records exist for this employee create one
    If (CR_snapFuture(EmpNo, "VAC") = 0) Or _
        (((UCase(glbCompEntVac$) = "M" Or UCase(glbCompEntVac$) = "N") And Not _
        (glbCompSerial = "S/N - 2355W" And rsEntRules("VE_MANUAL") <> 0)) And chkAnnualized.Value = False) Then
        
        snapFuture.AddNew
        snapFuture("EN_COMPNO") = "001"
        snapFuture("EN_EMPNBR") = EmpNo&
        snapFuture("EN_ENTSORT") = 1
        snapFuture("EN_TYPE_TABL") = "ENTT"
        snapFuture("EN_TYPE") = "VAC"
        snapFuture("EN_DESC") = "Vacation"
        'Ticket #19748
        'Get the first day of the month as per the From Date. If Effective Date is not same as the From Date (Future Entitl.)
        'then, for monthly entitlement, the From Date(Day) of the Month is same as the Effective Date (Day) in the
        'Report. e.g. Jan 15 - Jan 31, Feb 15 - Feb 28/29 - to avoid all this, I am comparing the Day part of the
        'As of Date to see if I have first day of the month date provided it matches From Date (Future Entitl.) on the screen.
        If Day(fglbAsOf) <> Day(dlpDateRange(0)) Then
            'snapFuture("EN_FDATE") = Format(month(fglbAsOf) & "/" & Day(dlpDateRange(0)) & "/" & Year(fglbAsOf), "mm/dd/yyyy")
            If Day(DateAdd("m", -1, fglbToDate) + 1) = 31 Then
                snapFuture("EN_FDATE") = DateAdd("m", -1, fglbToDate) + 2
            ElseIf Day(DateAdd("m", -1, fglbToDate) + 1) = 29 And month(fglbToDate) = 2 Then
                snapFuture("EN_FDATE") = CVDate(month(fglbToDate) & "/1/" & Year(fglbToDate))
            Else
                snapFuture("EN_FDATE") = DateAdd("m", -1, fglbToDate) + 1
            End If
        Else
            'snapFuture("EN_FDATE") = fglbAsOf
            If Day(DateAdd("m", -1, fglbToDate) + 1) = 31 Then
                snapFuture("EN_FDATE") = DateAdd("m", -1, fglbToDate) + 2
            ElseIf Day(DateAdd("m", -1, fglbToDate) + 1) = 29 And month(fglbToDate) = 2 Then
                snapFuture("EN_FDATE") = CVDate(month(fglbToDate) & "/1/" & Year(fglbToDate))
            Else
                snapFuture("EN_FDATE") = DateAdd("m", -1, fglbToDate) + 1
            End If
        End If
        snapFuture("EN_TDATE") = fglbToDate
        snapFuture("EN_SURNAME") = snapEntitle("ED_SURNAME")
        snapFuture("EN_FNAME") = snapEntitle("ED_FNAME")
        snapFuture("EN_WRKEMP") = glbUserID
        snapFuture("EN_LDATE") = Now
        snapFuture("EN_LTIME") = Time$
        snapFuture("EN_LUSER") = glbUserID
        snapFuture.Update
    End If

    If IsNull(snapFuture("EN_ENTITLE")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapFuture("EN_ENTITLE")
    End If

    'Ticket #22206 - Not to include Previous in the future entitlement calculation.
    'If IsNull(snapEntitle("ED_VAC")) Then
        dblPrevEntitle# = 0
    'Else
    '    dblPrevEntitle# = snapEntitle("ED_VAC")
    'End If

    spt = snapEntitle("ED_PT")

    If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec

    varStartDate = snapEntitle(fglbWDate$)

    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    rsJOB.Close
    If glbLinamar Then dblDHours# = 8

    xAsOf = fglbAsOf
    dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    intWhereFit& = -1

    If rsEntRules("VE_EMONTH") > 0 Then
        If dblServiceYears# >= CDbl(rsEntRules("VE_BMONTH")) And dblServiceYears# <= CDbl(rsEntRules("VE_EMONTH")) Then
            intWhereFit& = x%
            If Len(rsEntRules("VE_ENTITLE")) > 0 Then if_Entitle = True
            If Len(rsEntRules("VE_PCT")) > 0 Then if_Vacation = True
        End If
    End If

    'Ticket #16145 - Check for Mitchell Plastics their new hire and less than 12months Seniority logic
    'if true then call procedure to compute the entitlement for < 12 months logic
    'if new hire with Seniority between entitlement date then 0 entitlement
    'Then Goto Contd_Mitchell
    If glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") = 0 Then   'Mitchell Plastics
        If CVDate(varStartDate) >= CVDate(dlpDateRange(0)) And CVDate(varStartDate) <= CVDate(dlpDateRange(1)) Then
            if_Entitle = True
            dblEntitleUpd = 0
            GoTo Contd_Mitchell
        ElseIf dblServiceYears# < 12 And rsEntRules("VE_DIV") = "ULT" Then
            if_Entitle = True
            dblEntitleUpd = Assign_Entitlements_Mitchell(month(CVDate(varStartDate))) * dblDHours#
            GoTo Contd_Mitchell
        ElseIf dblServiceYears# < 12 And rsEntRules("VE_DIV") = "MIT" Then  ' 24 -> 12 'Ticket #23034 Franks 01/18/2012
            if_Entitle = True
            dblEntitleUpd = Assign_Entitlements_Mitchell_MIT(month(CVDate(varStartDate))) * dblDHours#
            GoTo Contd_Mitchell
        End If
    End If

    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges

    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
    ' which represents if Sick and Vacation entitlements
    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
    ' and read on system startup.

    ' In this routine we work independantly of SICK/VACATIon entitlement.
    '  fglbCompMonthly% - is the independant representation
        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
        'Procedure modUpdateSelection is used to set
        'fglbCompMonthly based on values it finds for global variables
        ' and what the user wants to manipulate (sick/Vac)

    'optD indicates if Entitlement entered is Daily or yearly based
    ' if daily then max entitlement is based on entitlement * hours they work.

    ' we have   Entitle = existing entitmenet (stored presently
    '           NewEntitle = amount entered onto screen = medentitle(index)
    '           EntitleUpd  = value to update record with

    If if_Entitle Then
        dblNewEntitle# = rsEntRules("VE_ENTITLE")
        dblNewMax# = 0
        If rsEntRules("VE_TYPE") = "D" Then           ' Entitlements entered in days
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If rsEntRules("VE_TYPE") = "F" Then
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX") * dblFTEHours# * dblDHours#
            dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            dblEntitleUpd = dblNewEntitle
        End If
        If rsEntRules("VE_TYPE") = "H" Then
            If rsEntRules("VE_MAX") <> 0 Then dblNewMax# = rsEntRules("VE_MAX")
        End If
        dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values

         If dblNewMax <> 0 Then          'only do if not zero
            If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                dblEntitleUpd = dblNewMax - dblPrevEntitle#
            End If
        End If

        DtTm = Now
    End If

'    If if_Vacation Then
'        If glbCBrant And Len(rsEntRules("VE_SECTION")) > 0 And snapEntitle("ED_SECTION") >= rsEntRules("VE_SECTION") Then
'            VacpcN = rsEntRules("VE_PCT") + dblEntitle#
'        Else
'            VacpcN = rsEntRules("VE_PCT")
'        End If
'        VacpcO = snapEntitle("ED_VACPC")
'        VED_DIV = snapEntitle("ED_DIV")
'        VED_PT = snapEntitle("ED_PT")
'        If IsNumeric(rsEntRules("VE_PCT")) Then snapEntitle("ED_VACPC") = rsEntRules("VE_PCT")
'
'    End If
    
Contd_Mitchell:
    If if_Entitle Then

        'If glbCompSerial = "S/N - 2188W" Then  'Ticket #8887
        '    dblEntitleUpd = Round(dblEntitleUpd, 0)
        If glbCompSerial = "S/N - 2297W" Then
            If dblEntitleUpd >= 14.9 And dblEntitleUpd <= 15.1 Then
                dblEntitleUpd = 15
            ElseIf dblEntitleUpd >= 19.9 And dblEntitleUpd <= 20.1 Then
                dblEntitleUpd = 20
            ElseIf dblEntitleUpd >= 25.1 And dblEntitleUpd <= 25.1 Then
                dblEntitleUpd = 25
            End If
        End If
        If glbCBrant And Len(rsEntRules("VE_SECTION")) > 0 Then
            dblEntitleUpd = rsEntRules("VE_PCT") + dblEntitle#
        End If


        If isLast And glbCompSerial = "S/N - 2376W" Then '#9536 on Oct 21,2005 George
            If dblDHours# <> 0 Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                dblEntitleDays = Round((dblEntitleDays / 0.25 + 0.1), 0) * 0.25 ' round to 1/4 days
                dblEntitleUpd = dblEntitleDays * dblDHours#
            End If
        End If

        'Hemu - 12/31/2003 End
        'Added by bryan 13/Jun/06 Ticket#10916
        snapFuture("EN_ENTITLE") = dblEntitleUpd
        snapFuture("EN_DHRS") = dblDHours#
    End If
    snapFuture.Update


lblNextRec:
gdbAdoIhr001.CommitTrans

snapFuture.Close

modAnnVacation = True

Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modAnnVac", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    'Rollback
    Resume Next
Else
    Unload Me
End If
End Function

Private Function CR_SnapHour() As Boolean

Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$, strTm$, x%

CR_SnapHour = False
On Error GoTo CR_SnapHourEntitle_Err
strTm$ = Time$
Dim Dt As Variant
Dt = Date$

Screen.MousePointer = HOURGLASS
SQLQ = "SELECT JH_DHRS,JH_FTENUM,ED_EMPNBR, ED_SURNAME, ED_FNAME,ED_DHRS,ED_DOH,ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1 "
If glbOracle Then
    SQLQ = SQLQ & "FROM HREMP, HR_JOB_HISTORY WHERE HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_CURRENT<>0"
Else
    SQLQ = SQLQ & "FROM HREMP INNER JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT<>0"
End If

SQLQ = SQLQ & " AND " & getWSQLQ_HR("") & " AND " & glbstrSelCri

If snapHourEntitle.State <> 0 Then snapHourEntitle.Close
snapHourEntitle.Open SQLQ, gdbAdoIhr001, adOpenStatic


Screen.MousePointer = DEFAULT
CR_SnapHour = True

Exit Function

CR_SnapHourEntitle_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapHourEntitle", "Entitlements/EMP", "Select")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function CR_snapFuture(EmpNo As Long, eType As String) As Long
    Dim SQLQ As String
    Dim lngRecs As Long

    SQLQ = "SELECT EN_COMPNO, EN_EMPNBR, EN_ID, EN_ENTSORT, EN_TYPE_TABL, EN_TYPE, EN_DESC, EN_FDATE, EN_TDATE, EN_ENTITLE, EN_DHRS, EN_TAKEN, "
    SQLQ = SQLQ & "EN_WRKEMP, EN_LDATE, EN_LTIME, EN_LUSER, EN_SURNAME, EN_FNAME From HR_NEXTENT_WRK "
    SQLQ = SQLQ & "WHERE EN_EMPNBR=" & CStr(EmpNo) & " AND EN_TYPE='" & eType & "' AND EN_WRKEMP = '" & glbUserID & "'"
    snapFuture.Open SQLQ, gdbAdoIhr001W, adOpenDynamic, adLockOptimistic, adCmdText

    lngRecs = snapFuture.RecordCount
    CR_snapFuture = lngRecs
End Function

Private Function getcMaster() As Boolean
    Dim SQLQ As String
    Dim rs As New ADODB.Recordset
    Dim retVal As Boolean
    
    retVal = False
    
    SQLQ = "SELECT PC_VACENT, PC_SICKENT FROM HRPARCO"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockReadOnly, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If rs("PC_VACENT") = "M" Or rs("PC_SICKENT") = "M" Then
            retVal = True
        End If
    End If
    rs.Close
    
    getcMaster = retVal
    
End Function

Public Function getEOM(Mnt As Variant) As Integer
   Dim myDate As Date
   Dim myMonth As String
   
   Dim NextMonth As Date, EndOfMonth As Date
   If IsNumeric(Mnt) Then
        myMonth = MonthName(Mnt, True)
    Else
        myMonth = Mnt
    End If
   
   myDate = Format("1/" & myMonth & "/2005", "dd/mmm/yyyy")
   NextMonth = DateAdd("m", 1, myDate)
   EndOfMonth = NextMonth - DatePart("d", NextMonth)
   getEOM = Day(EndOfMonth)

End Function

Private Function Assign_Entitlements_Mitchell(xMonth)
    
    'New Logic - Ticket #15130 - Paid for logic - # of Days based on the month of hire
    Select Case xMonth
        Case 7: Assign_Entitlements_Mitchell = 10
        Case 1: Assign_Entitlements_Mitchell = 5
        Case 8: Assign_Entitlements_Mitchell = 9
        Case 2: Assign_Entitlements_Mitchell = 4
        Case 9: Assign_Entitlements_Mitchell = 8
        Case 3: Assign_Entitlements_Mitchell = 3
        Case 10: Assign_Entitlements_Mitchell = 7
        Case 4: Assign_Entitlements_Mitchell = 3
        Case 11: Assign_Entitlements_Mitchell = 7
        Case 5: Assign_Entitlements_Mitchell = 2
        Case 12: Assign_Entitlements_Mitchell = 6
        Case 6: Assign_Entitlements_Mitchell = 1
    End Select

End Function

Private Function Assign_Entitlements_Mitchell_MIT(xMonth)
    'New Logic for Mitchell Division - Ticket #18124 - # of Days based on the month of hire
    Select Case xMonth
        Case 7: Assign_Entitlements_Mitchell_MIT = 5
        Case 1: Assign_Entitlements_Mitchell_MIT = 10
        Case 8: Assign_Entitlements_Mitchell_MIT = 4
        Case 2: Assign_Entitlements_Mitchell_MIT = 9
        Case 9: Assign_Entitlements_Mitchell_MIT = 3
        Case 3: Assign_Entitlements_Mitchell_MIT = 8
        Case 10: Assign_Entitlements_Mitchell_MIT = 3
        Case 4: Assign_Entitlements_Mitchell_MIT = 7
        Case 11: Assign_Entitlements_Mitchell_MIT = 2
        Case 5: Assign_Entitlements_Mitchell_MIT = 7
        Case 12: Assign_Entitlements_Mitchell_MIT = 1
        Case 6: Assign_Entitlements_Mitchell_MIT = 6
    End Select

End Function

