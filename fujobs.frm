VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUJobs 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Position Mass Update"
   ClientHeight    =   7695
   ClientLeft      =   945
   ClientTop       =   1650
   ClientWidth     =   11040
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7695
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkUpdAttendance 
      Caption         =   "Update Employee's Attendance with New Salary"
      Height          =   255
      Left            =   1440
      TabIndex        =   16
      Tag             =   "40-Update Attendance records with Salary -y/n"
      Top             =   6480
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.ComboBox cmbPrecision 
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
      Left            =   8400
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Enter Number of Decimal Places for Salary"
      Top             =   3255
      Visible         =   0   'False
      Width           =   870
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Position Code"
      Top             =   1440
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   0
      LookupType      =   5
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   2
      Tag             =   "00-Enter Position Status"
      Top             =   1080
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "JBST"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Enter Union Code"
      Top             =   720
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
      Index           =   1
      Left            =   1800
      TabIndex        =   0
      Tag             =   "00-Group - Code"
      Top             =   360
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "JBGC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.CheckBox chkCompa 
      Caption         =   "Recalculate Compa-Ratio"
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
      Left            =   600
      TabIndex        =   10
      Tag             =   "40-Recalculate Compa-Ratio -y/n"
      Top             =   4140
      Width           =   2655
   End
   Begin VB.ComboBox cmbRound 
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
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "Select Yes / No"
      Top             =   3255
      Width           =   735
   End
   Begin Threed.SSOption optDollars 
      Height          =   225
      Left            =   600
      TabIndex        =   8
      Tag             =   "Increase/Decrease entered in dollars"
      Top             =   3780
      Width           =   2640
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "   Dollar Increase/Decrease"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSOption optPct 
      Height          =   225
      Left            =   3480
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "Increase/Decrease entered as a percent (enter with decimal)"
      Top             =   3780
      Width           =   2955
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   "   Percentage Increase/Decrease"
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
   Begin MSMask.MaskEdBox medChng 
      Height          =   300
      Left            =   2880
      TabIndex        =   5
      Tag             =   "11-Amount to change by"
      Top             =   3255
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
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
   Begin Threed.SSPanel SSPSal 
      Height          =   1335
      Left            =   1320
      TabIndex        =   26
      Top             =   4860
      Visible         =   0   'False
      Width           =   6135
      _Version        =   65536
      _ExtentX        =   10821
      _ExtentY        =   2355
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
      BevelOuter      =   0
      Begin INFOHR_Controls.DateLookup dlpEDate 
         Height          =   285
         Left            =   1680
         TabIndex        =   13
         Tag             =   "41-Enter Salary Effective Date"
         Top             =   360
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin Threed.SSOption optAmount 
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Tag             =   "Update Salary by the Amount to Change By"
         Top             =   720
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Update Salary by the Amount to Change By"
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
      Begin Threed.SSOption optStep 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Tag             =   "Update Salary to Step Amount"
         Top             =   1080
         Width           =   3855
         _Version        =   65536
         _ExtentX        =   6800
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Update Salary to Step Amount"
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
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   12
         Tag             =   "01-Reason code "
         Top             =   30
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "SDRC"
      End
      Begin VB.Label lblEDate 
         Caption         =   "Effective Date"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblReason 
         Caption         =   "Reason"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   1095
      End
   End
   Begin INFOHR_Controls.CodeLookup clpGrid 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "JBGD"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   3300
      TabIndex        =   17
      Top             =   7320
      Visible         =   0   'False
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDEM"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.CheckBox chkSalary 
      Caption         =   "Update Employee Salary"
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
      Left            =   600
      TabIndex        =   11
      Tag             =   "40-Update Employee Salary -y/n"
      Top             =   4500
      Width           =   2775
   End
   Begin VB.Label lblSelCri2 
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
      Left            =   1440
      TabIndex        =   32
      Top             =   6960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblEStatus 
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
      Left            =   1800
      TabIndex        =   31
      Top             =   7350
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblDecimal 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Decimal Precision"
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
      Left            =   7080
      TabIndex        =   30
      Top             =   3315
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblGrid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Grid Category"
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
      Left            =   600
      TabIndex        =   29
      Top             =   1860
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblSelCri 
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
      Left            =   360
      TabIndex        =   25
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   600
      TabIndex        =   24
      Top             =   1470
      Width           =   975
   End
   Begin VB.Label lblPosStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Position Status"
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
      Left            =   600
      TabIndex        =   23
      Top             =   1110
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Amount to Change by:"
      Height          =   195
      Left            =   600
      TabIndex        =   22
      Top             =   3300
      Width           =   1890
   End
   Begin VB.Label lblPosition 
      Caption         =   "Position Scale Update"
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
      Left            =   360
      TabIndex        =   21
      Top             =   2940
      Width           =   2055
   End
   Begin VB.Label lblRound 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Round"
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
      Left            =   5520
      TabIndex        =   20
      Top             =   3308
      Width           =   615
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Union"
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
      Left            =   600
      TabIndex        =   19
      Top             =   750
      Width           =   1140
   End
   Begin VB.Label lblPosGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   600
      TabIndex        =   18
      Top             =   420
      Width           =   1035
   End
End
Attribute VB_Name = "frmUJobs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dynSH_Job1 As New ADODB.Recordset
Dim SQLQGen
Dim fglbCOMPA#, fglbGRADE$
Dim MsgSal, IfDisplay
Dim OSalary, NSalary, OEDate, NEDate, ONDate, NNDate, EmpNo&, dblWHours#, OTOTAL
Dim oPayP, NPayp, OJOB1, OSalCD
Dim oStep
Dim oGrid, oPayrollID
Dim lngRecs&, UpdRecs&
Dim WSQLQ
Dim MailBody
Dim fglbDhrs

Private Function AUDITSALY()
Dim TA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim TB As New ADODB.Recordset
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDITSALY = False


TB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & EmpNo&, gdbAdoIhr001, adOpenKeyset
If Not TB.EOF Then
    If IsNull(TB("ED_PT")) Then
        xPT = ""
    Else
        xPT = TB("ED_PT")
    End If
    If IsNull(TB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = TB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
TB.Close
'TA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'strFields added by Bryan on 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_SALARY, AU_OLDSAL, AU_PAYP, AU_OLDPAYP, AU_JOB, AU_GRID, "
strFields = strFields & "AU_PAYROLL_ID, AU_SALCD, AU_WHRS, AU_SEDATE, AU_SNDATE, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, "
strFields = strFields & "AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_SREASON "
TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

If OSalary <> NSalary Then GoTo MODUPD
If OEDate <> NEDate Then GoTo MODUPD
'If ONDate <> NNDate Then GoTo MODUPD
GoTo MODNOUPD

MODUPD:
TA.AddNew
TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR"
TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL"
TA("AU_EARN_TABL") = "EARN"
TA("AU_NEWEMP") = "N"
TA("AU_PTUPL") = xPT
TA("AU_DIVUPL") = xDiv

TA("AU_SALARY") = NSalary
TA("AU_OLDSAL") = OSalary
TA("AU_PAYP") = oPayP ' FRANK 4/5/2000    'NPayp  Laura jan 28, 1998
TA("AU_OLDPAYP") = oPayP    '    ""
TA("AU_JOB") = OJOB1         ' FRANK 4/5/2000
TA("AU_GRID") = oGrid
If glbMulti Then TA("AU_PAYROLL_ID") = oPayrollID
TA("AU_SALCD") = OSalCD
TA("AU_WHRS") = dblWHours# 'ADDED BY RAUBREY 7/7/97
If OEDate <> NEDate Then TA("AU_SEDATE") = IIf(IsDate(NEDate), NEDate, Null)   'Jaddy 11/15/99
If ONDate <> NNDate Then TA("AU_SNDATE") = IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99

'Ticket #23666 - Update with Salary Reason for Change as well.
TA("AU_SREASON") = clpCode(4).Text

TA("AU_COMPNO") = "001"
TA("AU_EMPNBR") = EmpNo&

'Ticket #23943 - Town of Orangeville noticed the LDATE was not getting updated properly - Jerry asked to fix this as per Salary screen.
If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
    TA("AU_LDATE") = Format(DateAdd("d", 14, NEDate), "SHORT DATE")
Else
    'Ticket #23943 - Town of Orangeville
    If glbCompSerial = "S/N - 2383W" Then
        If CVDate(NEDate) > CVDate(Date) Then
            TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
        Else
            TA("AU_LDATE") = Date
        End If
    Else
        TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
    End If
End If
'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")

TA("AU_LUSER") = glbUserID
TA("AU_LTIME") = Time$
TA("AU_UPLOAD") = "N"
TA("AU_TYPE") = "A"
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & EmpNo&
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then TA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If
TA.Update

MODNOUPD:
AUDITSALY = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Function chkUSelect()
Dim Msg$, DgDef As Variant, Response%
chkUSelect = False

If Not clpCode(1).ListChecker Then Exit Function    'Release 8.0
'If Len(clpCode(1).Text) > 0 Then
'    If clpCode(1).Caption = "Unassigned" Then
'        MsgBox lStr("Position Group must be valid")
'        clpCode(1).SetFocus
'        Exit Function
'    End If
'End If

If Not clpCode(2).ListChecker Then Exit Function    'Release 8.0
'If Len(clpCode(2).Text) > 0 Then      'laura 03/26/98
'    If clpCode(2).Caption = "Unassigned" Then
'        MsgBox lStr("Union Code must be valid")
'        clpCode(2).SetFocus
'        Exit Function
'    End If
'End If

If Not clpCode(3).ListChecker Then Exit Function    'Release 8.0

If Not clpJob.ListChecker Then Exit Function    'Release 8.0

If Not clpGrid.ListChecker Then Exit Function    'Release 8.0

If Len(medChng) < 1 Then
    MsgBox "Amount to change by is a required field"
    medChng.SetFocus
    Exit Function
End If

If Not IsNumeric(medChng) Then
    MsgBox "Change amount must be numeric"
    medChng.SetFocus
    Exit Function
End If

'If optPct And (medchng > 1 Or medchng < 1) Then
'    Msg$ = "Are you sure you wish to update the "
'    Msg$ = Msg$ & Chr(10) & "Jobs with this group by "
'     Msg$ = Msg$ & Chr(10) & CStr(medchng) & " PERCENT?"
'    'Msg$ = Msg$ & Chr(10) & CStr(medChng * 100) & " PERCENT?"'changed by RAUBREY 6/3/97
'    dgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
'    Response% = MsgBox(Msg$, dgDef, "Warning!")
'    If Response% = IDNO Then
'        Exit Function
'    End If
'End If
If chkSalary Then 'FRANK 5/24/2000
    If Len(clpCode(4).Text) < 1 Then
        MsgBox "Reason for Salary Change must be entered"
        clpCode(4).SetFocus
        Exit Function
    Else
        If clpCode(4).Caption = "Unassigned" Then
            MsgBox "Reason for Salary Change must be valid"
            clpCode(4).SetFocus
            Exit Function
        End If
    End If
End If
If chkSalary Then 'FRANK 4/12/2000
    If Len(dlpEDate.Text) < 1 Then
        MsgBox "Effective Date must be entered"
        dlpEDate.SetFocus
        Exit Function
    End If
    If Not IsDate(dlpEDate.Text) Then
        MsgBox "Effective Date is not a valid date"
        dlpEDate.SetFocus
        Exit Function
    End If
End If

'Ticket #28969 - Employment Status
If chkSalary Then
    If Not clpCode(0).ListChecker Then Exit Function
End If

chkUSelect = True

End Function

Private Sub chkCompa_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkSalary_Click() 'Frank 4/14/2000
If chkSalary.Value Then
    SSPSal.Visible = True
    chkUpdAttendance.Visible = True
    
    'Ticket #28969 - Employment Status
    lblSelCri2.Visible = True
    lblEStatus.Visible = True
    clpCode(0).Visible = True
Else
    SSPSal.Visible = False
    chkUpdAttendance.Visible = False

    'Ticket #28969 - Employment Status
    lblSelCri2.Visible = False
    lblEStatus.Visible = False
    clpCode(0).Visible = False
End If
End Sub

Private Sub chkSalary_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkUpdAttendance_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbPrecision_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbRound_Click()
    If cmbRound.Text = "Yes" Then
        cmbPrecision.Visible = True
        lblDecimal.Visible = True
        cmbPrecision.Text = glbCompDecHR
    Else
        cmbPrecision.Visible = False
        lblDecimal.Visible = False
    End If
End Sub

Private Sub cmbRound_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdModify_Click()
If chkUSelect() Then
    Call modJobUpd
End If
End Sub


Private Sub Upd_Related_Data()
Dim SQLQ As String, Msg As String
Dim dynSH_Job As New ADODB.Recordset
Dim ssalary, MidSalary, Compa
Dim SalaryGrade$
'On Error GoTo UpRel_Err
If glbOracle Then
    If glbMultiGrid Then
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY,HR_JOB_HISTORY,HRJOB_GRADE"
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HR_JOB_HISTORY.JH_GRID "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HRJOB_GRADE.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HRJOB_GRADE.JB_GRID "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    Else
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB.* "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY,HR_JOB_HISTORY,HRJOB"
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    End If
Else
    If glbMultiGrid Then
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM (HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HR_JOB_HISTORY.JH_GRID "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        SQLQ = SQLQ & " INNER JOIN HRJOB_GRADE "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HRJOB_GRADE.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HRJOB_GRADE.JB_GRID "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    Else
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB.* "
        SQLQ = SQLQ & " FROM (HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        SQLQ = SQLQ & " INNER JOIN HRJOB"
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    End If
End If
SQLQ = SQLQ & " AND " & WSQLQ
SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")"

If glbNoNONE And glbNoEXEC Then   'Hemu -EXE
    SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG <> 'NONE' AND ED_ORG <> 'EXEC') "
ElseIf glbNoNONE Then
    SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG <> 'NONE') "
ElseIf glbNoEXEC Then   'Hemu -EXE
    SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG <> 'EXEC') " 'Hemu -EXE
End If
dynSH_Job.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
With dynSH_Job
    Do Until .EOF
        ssalary = !SH_SALARY
        MidSalary = 0
        SalaryGrade$ = !JB_CODE
        If IsNumeric(!JB_MIDPOINT) Then MidSalary = dynSH_Job("JB_S" & !JB_MIDPOINT)
        'D added by Bryan 28/Sep/05 Ticket#9354
        If !SH_SALCD = "H" And !JB_SALCD = "A" Then ssalary = ssalary * !SH_WHRS * 52
        If !SH_SALCD = "M" And !JB_SALCD = "A" Then ssalary = (ssalary * !SH_WHRS * 52) / 12
        If !SH_SALCD = "D" And !JB_SALCD = "A" Then
            If GetLeapYear(Year(Date)) Then
                ssalary = ssalary * 366
            Else
                ssalary = ssalary * 365
            End If
            
            'Ticket #17654 - formula correction
            ssalary = ssalary / (!SH_WHRS * 52) * GetJHData(!SH_EMPNBR, "JH_DHRS", 0)
        End If
        
        If !SH_SALCD = "A" And !JB_SALCD = "H" Then MidSalary = MidSalary * !SH_WHRS * 52
        If !SH_SALCD = "M" And !JB_SALCD = "H" Then MidSalary = (MidSalary * !SH_WHRS * 52) / 12
        If !SH_SALCD = "D" And !JB_SALCD = "H" Then
            If GetLeapYear(Year(Date)) Then
                MidSalary = (MidSalary * !SH_WHRS * 52) / 366
            Else
                MidSalary = (MidSalary * !SH_WHRS * 52) / 366
            End If
            
            'Ticket #17654 - formula correction
            MidSalary = MidSalary * GetJHData(!SH_EMPNBR, "JH_DHRS", 0)
        End If
        
        Compa = 0
        If MidSalary <> 0 Then Compa = (ssalary / MidSalary) * 100
        If Compa > 999.99 Then Compa = 999.99
        !SH_COMPA = Compa
        !SH_TRANSDATE = Now
        !SH_LDATE = Now
        !SH_LTIME = Time$
        !SH_LUSER = glbUserID 'glbLEE_ID
        .Update
        .MoveNext
    Loop
End With
Exit Sub

UpRel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Error", "HRJobs", "Update")

Resume Next

End Sub

Private Sub GetNewStepSalary(dblNewSalary) 'Frank 4/12/2000
Dim x%, cX$, xSalGrade, SQLQ
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#
Dim snapJob As New ADODB.Recordset
Dim xGrid2 As Boolean
    
    xGrid2 = False
    
    'SET COMPA RATIO
    '================
    If glbMultiGrid Then
        SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "' AND JB_GRID='" & dynSH_Job1("SH_GRID") & "'"
    Else
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "'"
    End If
    
    snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ssalary@ = dblNewSalary
    dblHoursPerWeek# = dynSH_Job1("SH_WHRS")

    fglbGRADE$ = "00"
    xSalGrade = dblNewSalary
    fglbDhrs = GetJHData(dynSH_Job1("SH_EMPNBR"), "JH_DHRS", 0)
    
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For x% = 1 To 20
        If IsNumeric(dynSH_Job1("JB_S" & Format(x%, "##"))) Then
            If snapJob("JB_SALCD") = "H" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    xSalGrade = (snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)) / 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)
                    If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                        xSalGrade = snapJob("JB_S" & Format(x%, "##") & "A")
                        xGrid2 = True
                    End If
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52) / 366
                    Else
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52) / 365
                    End If
                    
                    'Ticket #17654 - formula correction
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * fglbDhrs
                End If
            ElseIf snapJob("JB_SALCD") = "A" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    If dblHoursPerWeek# = 0 Then
                        xSalGrade = 0
                    Else
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52)
                    End If
                    If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                        xSalGrade = snapJob("JB_S" & Format(x%, "##") & "A")
                        xGrid2 = True
                    End If
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) / 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 366
                    Else
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 365
                    End If
                    
                    'Ticket #17654 - formula correction
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52) * fglbDhrs
                End If
            End If
            If glbCompSerial = "S/N - 2378W" And xGrid2 = True Then 'Town of Aurora
                If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(x%, "##") & "A") > 0 Then
                    cX$ = CStr(x)
                    If x% <= 9 Then cX$ = "0" & cX$
                    fglbGRADE$ = cX$
                End If
            Else
                If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(x%, "##")) > 0 Then
                    cX$ = CStr(x)
                    If x% <= 9 Then cX$ = "0" & cX$
                    fglbGRADE$ = cX$
                End If
            End If
        End If
    Next x%
    snapJob.Close
End Sub

Private Function GetNewStepSalary_1(dblNewSalary)
Dim x%, cX$, xSalGrade, SQLQ
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#
Dim snapJob As New ADODB.Recordset
Dim xGrid2 As Boolean
Dim xStep As Integer
    
    xGrid2 = False
    
    'SET COMPA RATIO
    '================
    If glbMultiGrid Then
        SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "' AND JB_GRID='" & dynSH_Job1("SH_GRID") & "'"
    Else
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "'"
    End If
    
    snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ssalary@ = dblNewSalary
    dblHoursPerWeek# = dynSH_Job1("SH_WHRS")
    fglbDhrs = GetJHData(dynSH_Job1("SH_EMPNBR"), "JH_DHRS", 0)
    
    fglbGRADE$ = "00"
    xSalGrade = dblNewSalary
    
    If Len(oStep) > 0 Then
        xStep = CInt(oStep)
    Else
        xStep = 0
    End If
    
    If xStep <> 0 Then
        If IsNumeric(dynSH_Job1("JB_S" & Format(xStep, "##"))) Then
            If snapJob("JB_SALCD") = "H" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    dblNewSalary = (snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52)) / 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52)
                    If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                        dblNewSalary = snapJob("JB_S" & Format(xStep, "##") & "A")
                        xGrid2 = True
                    End If
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52) / 366
                    Else
                        dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52) / 365
                    End If
                    
                    'Ticket #17654 - formula correction
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * fglbDhrs
                End If
            ElseIf snapJob("JB_SALCD") = "A" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    If dblHoursPerWeek# = 0 Then
                        dblNewSalary = 0
                    Else
                        dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) / (dblHoursPerWeek# * 52)
                    End If
                    If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                        dblNewSalary = snapJob("JB_S" & Format(xStep, "##") & "A")
                        xGrid2 = True
                    End If
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) / 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * 366
                    Else
                        dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * 365
                    End If
                    
                    'Ticket #17654 - formula correction
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) / (dblHoursPerWeek# * 52) * fglbDhrs
                End If
            End If
            If glbCompSerial = "S/N - 2378W" And xGrid2 = True Then 'Town of Aurora
                If dynSH_Job1("JB_S" & Format(xStep, "##") & "A") > 0 Then
                    cX$ = CStr(xStep)
                    If x% <= 9 Then cX$ = "0" & cX$
                    fglbGRADE$ = cX$
                End If
            Else
                If dynSH_Job1("JB_S" & Format(xStep, "##")) > 0 Then
                    cX$ = CStr(xStep)
                    If x% <= 9 Then cX$ = "0" & cX$
                    fglbGRADE$ = cX$
                End If
            End If
        End If
    Else
        'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
        'For X% = 1 To 11
        'For X% = 1 To 15
        For x% = 1 To 20
            If IsNumeric(dynSH_Job1("JB_S" & Format(x%, "##"))) Then
                If snapJob("JB_SALCD") = "H" Then
                    If dynSH_Job1("SH_SALCD") = "H" Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##"))
                    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                        xSalGrade = (snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)) / 12
                    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)
                        If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                            xSalGrade = snapJob("JB_S" & Format(x%, "##") & "A")
                            xGrid2 = True
                        End If
                    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                        If GetLeapYear(Year(Date)) Then
                            xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52) / 366
                        Else
                            xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52) / 365
                        End If
                    
                        'Ticket #17654 - formula correction
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * fglbDhrs
                    End If
                ElseIf snapJob("JB_SALCD") = "A" Then
                    If dynSH_Job1("SH_SALCD") = "H" Then
                        If dblHoursPerWeek# = 0 Then
                            xSalGrade = 0
                        Else
                            xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52)
                        End If
                        If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                            xSalGrade = snapJob("JB_S" & Format(x%, "##") & "A")
                            xGrid2 = True
                        End If
                    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) / 12
                    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##"))
                    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                        If GetLeapYear(Year(Date)) Then
                            xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 366
                        Else
                            xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 365
                        End If
                        
                        'Ticket #17654 - formula correction
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52) * fglbDhrs
                    End If
                End If
                If glbCompSerial = "S/N - 2378W" And xGrid2 = True Then 'Town of Aurora
                    If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(x%, "##") & "A") > 0 Then
                        cX$ = CStr(x)
                        If x% <= 9 Then cX$ = "0" & cX$
                        fglbGRADE$ = cX$
                    End If
                Else
                    If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(x%, "##")) > 0 Then
                        cX$ = CStr(x)
                        If x% <= 9 Then cX$ = "0" & cX$
                        fglbGRADE$ = cX$
                    End If
                End If
            End If
        Next x%
    End If
    
    GetNewStepSalary_1 = dblNewSalary
    snapJob.Close
End Function

Private Sub modSetCOMPA_GRADE_1(ByRef dblNewSalary)

Dim x%, cX$, xSalGrade, SQLQ
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#
Dim snapJob As New ADODB.Recordset
Dim xStep As Integer
Dim xGrid2 As Boolean

'SET COMPA RATIO
'================
If glbMultiGrid Then
    SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "' AND JB_GRID='" & dynSH_Job1("SH_GRID") & "'"
Else
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "'"
End If
snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
ssalary@ = dblNewSalary
dblHoursPerWeek# = dynSH_Job1("SH_WHRS")
fglbDhrs = GetJHData(dynSH_Job1("SH_EMPNBR"), "JH_DHRS", 0)

If Len(oStep) > 0 Then
    xStep = CInt(oStep)
Else
    xStep = 0
End If

If optStep Then
    If snapJob("JB_SALCD") = "H" Then
        If dynSH_Job1("SH_SALCD") = "H" Then
            If xStep <> 0 Then
                dblSsalary# = snapJob("JB_S" & Format(xStep, "##"))
            Else
                dblSsalary# = 0
            End If
        ElseIf dynSH_Job1("SH_SALCD") = "M" Then
            If dblHoursPerWeek# = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = (dblNewSalary * 12) / (dblHoursPerWeek# * 52)
            End If
        ElseIf dynSH_Job1("SH_SALCD") = "A" Then
            If dblHoursPerWeek# = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = dblNewSalary / (dblHoursPerWeek# * 52)
            End If
        ElseIf dynSH_Job1("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                dblSsalary# = (dblNewSalary * 366) / (dblHoursPerWeek# * 52)
            Else
                dblSsalary# = (dblNewSalary * 365) / (dblHoursPerWeek# * 52)
            End If
            
            'Ticket #17654 - formula correction
            If fglbDhrs = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = dblNewSalary / fglbDhrs
            End If
        End If
    ElseIf snapJob("JB_SALCD") = "A" Then
        If dynSH_Job1("SH_SALCD") = "H" Then
            dblSsalary# = (dblNewSalary * dblHoursPerWeek#) * 52
        ElseIf dynSH_Job1("SH_SALCD") = "M" Then
            dblSsalary# = dblNewSalary * 12
        ElseIf dynSH_Job1("SH_SALCD") = "A" Then
            If xStep <> 0 Then
                dblSsalary# = snapJob("JB_S" & Format(xStep, "##"))
            Else
                dblSsalary# = 0
            End If
        ElseIf dynSH_Job1("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                dblSsalary# = (dblNewSalary * 366)
            Else
                dblSsalary# = (dblNewSalary * 365)
            End If
            
            'Ticket #17654 - formula correction
            If fglbDhrs = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = (dblNewSalary / fglbDhrs) * dblHoursPerWeek# * 52
            End If
        End If
    End If
Else
    If snapJob("JB_SALCD") = "H" Then
        If dynSH_Job1("SH_SALCD") = "H" Then
            dblSsalary# = dblNewSalary
        ElseIf dynSH_Job1("SH_SALCD") = "M" Then
            If dblHoursPerWeek# = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = (dblNewSalary * 12) / (dblHoursPerWeek# * 52)
            End If
        ElseIf dynSH_Job1("SH_SALCD") = "A" Then
            If dblHoursPerWeek# = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = dblNewSalary / (dblHoursPerWeek# * 52)
            End If
        ElseIf dynSH_Job1("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                dblSsalary# = (dblNewSalary * 366) / (dblHoursPerWeek# * 52)
            Else
                dblSsalary# = (dblNewSalary * 365) / (dblHoursPerWeek# * 52)
            End If
            
            'Ticket #17654 - formula correction
            If fglbDhrs = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = dblNewSalary / fglbDhrs
            End If
        End If
    ElseIf snapJob("JB_SALCD") = "A" Then
        If dynSH_Job1("SH_SALCD") = "H" Then
            dblSsalary# = (dblNewSalary * dblHoursPerWeek#) * 52
        ElseIf dynSH_Job1("SH_SALCD") = "M" Then
            dblSsalary# = dblNewSalary * 12
        ElseIf dynSH_Job1("SH_SALCD") = "A" Then
            dblSsalary# = dblNewSalary
        ElseIf dynSH_Job1("SH_SALCD") = "D" Then
            If GetLeapYear(Year(Date)) Then
                dblSsalary# = (dblNewSalary * 366)
            Else
                dblSsalary# = (dblNewSalary * 365)
            End If
            
            'Ticket #17654 - formula correction
            If fglbDhrs = 0 Then
                dblSsalary# = 0
            Else
                dblSsalary# = (dblNewSalary / fglbDhrs) * dblHoursPerWeek# * 52
            End If
        End If
    End If
End If

' dkostka - 02/18/2002 - Added Val(Format(x, "@")) around expression to replace null with 0.
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 11 Then
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 15 Then
If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 20 Then
    Jb_No = Val(Format(dynSH_Job1("JB_S" & dynSH_Job1("JB_MIDPOINT")), "@"))
End If

fglbCOMPA# = 0

If Jb_No <> 0 And dblSsalary# <> 0 Then 'laura 03/23/98
  fglbCOMPA# = (dblSsalary# / Jb_No) * 100
End If

 
If fglbCOMPA# > 999.99 Then fglbCOMPA# = 999.99


fglbGRADE$ = "00"
xSalGrade = dblNewSalary

If xStep <> 0 Then
    If IsNumeric(dynSH_Job1("JB_S" & Format(xStep, "##"))) Then
        If snapJob("JB_SALCD") = "H" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                dblNewSalary = (snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52)) / 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52)
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##") & "A")
                    xGrid2 = True
                End If
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52) / 366
                Else
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * (dblHoursPerWeek# * 52) / 365
                End If
                
                'Ticket #17654 - formula correction
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * fglbDhrs
            End If
        ElseIf snapJob("JB_SALCD") = "A" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                If dblHoursPerWeek# = 0 Then
                    dblNewSalary = 0
                Else
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) / (dblHoursPerWeek# * 52)
                End If
                If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##") & "A")
                    xGrid2 = True
                End If
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) / 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                dblNewSalary = snapJob("JB_S" & Format(xStep, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * 366
                Else
                    dblNewSalary = snapJob("JB_S" & Format(xStep, "##")) * 365
                End If
                
                'Ticket #17654 - formula correction
                dblNewSalary = (snapJob("JB_S" & Format(xStep, "##")) / (dblHoursPerWeek# * 52)) * fglbDhrs
            End If
        End If
        If glbCompSerial = "S/N - 2378W" And xGrid2 = True Then 'Town of Aurora
            If dynSH_Job1("JB_S" & Format(xStep, "##") & "A") > 0 Then
                cX$ = CStr(xStep)
                If x% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        Else
            If dynSH_Job1("JB_S" & Format(xStep, "##")) > 0 Then
                cX$ = CStr(xStep)
                If x% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        End If
    End If
Else
    'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
    'For X% = 1 To 11
    'For X% = 1 To 15
    For x% = 1 To 20
        'Added D by Bryan 28/Sep/05 Ticket#9354
        If IsNumeric(dynSH_Job1("JB_S" & Format(x%, "##"))) Then
            If snapJob("JB_SALCD") = "H" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    xSalGrade = (snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)) / 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 366 / (dblHoursPerWeek# * 52)
                    Else
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 365 / (dblHoursPerWeek# * 52)
                    End If
                
                    'Ticket #17654 - formula correction
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * fglbDhrs
                End If
            ElseIf snapJob("JB_SALCD") = "A" Then
                If dynSH_Job1("SH_SALCD") = "H" Then
                    If dblHoursPerWeek# = 0 Then
                        xSalGrade = 0
                    Else
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52)
                    End If
                ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 12
                ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##"))
                ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 366
                    Else
                        xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 365
                    End If
                    
                    'Ticket #17654 - formula correction
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52) * fglbDhrs
                End If
            End If
            If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(x%, "##")) > 0 Then
                cX$ = CStr(x)
                If x% <= 9 Then cX$ = "0" & cX$
                fglbGRADE$ = cX$
            End If
        End If
    Next x%
End If

If IsNumeric(dynSH_Job1("JB_S1")) Then
    If dblSsalary# < dynSH_Job1("JB_S1") Then
        fglbGRADE$ = "00"
    End If
End If
snapJob.Close
End Sub


Private Sub modSetCOMPA_GRADE(dblNewSalary)

Dim x%, cX$, xSalGrade, SQLQ
Dim dblSsalary#, dblHoursPerWeek#, ssalary@
Dim Jb_No#
Dim snapJob As New ADODB.Recordset
'SET COMPA RATIO
'================
If glbMultiGrid Then
    SQLQ = "SELECT * FROM HRJOB_GRADE WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "' AND JB_GRID='" & dynSH_Job1("SH_GRID") & "'"
Else
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE='" & dynSH_Job1("SH_JOB") & "'"
End If
snapJob.Open SQLQ, gdbAdoIhr001, adOpenStatic

ssalary@ = dblNewSalary
dblHoursPerWeek# = dynSH_Job1("SH_WHRS")
fglbDhrs = GetJHData(dynSH_Job1("SH_EMPNBR"), "JH_DHRS", 0)

If snapJob("JB_SALCD") = "H" Then
    If dynSH_Job1("SH_SALCD") = "H" Then
        dblSsalary# = dblNewSalary
    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary * 12) / (dblHoursPerWeek# * 52)
        End If
    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
        If dblHoursPerWeek# = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = dblNewSalary / (dblHoursPerWeek# * 52)
        End If
    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = (dblNewSalary * 366) / (dblHoursPerWeek# * 52)
        Else
            dblSsalary# = (dblNewSalary * 365) / (dblHoursPerWeek# * 52)
        End If
        
        'Ticket #17654 - formula corrections
        If fglbDhrs = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = dblNewSalary / fglbDhrs
        End If
    End If
ElseIf snapJob("JB_SALCD") = "A" Then
    If dynSH_Job1("SH_SALCD") = "H" Then
        dblSsalary# = (dblNewSalary * dblHoursPerWeek#) * 52
    ElseIf dynSH_Job1("SH_SALCD") = "M" Then
        dblSsalary# = dblNewSalary * 12
    ElseIf dynSH_Job1("SH_SALCD") = "A" Then
        dblSsalary# = dblNewSalary
    ElseIf dynSH_Job1("SH_SALCD") = "D" Then
        If GetLeapYear(Year(Date)) Then
            dblSsalary# = (dblNewSalary * 366)
        Else
            dblSsalary# = (dblNewSalary * 365)
        End If
        
        'Ticket #17654 - formula corrections
        If fglbDhrs = 0 Then
            dblSsalary# = 0
        Else
            dblSsalary# = (dblNewSalary / fglbDhrs) * dblHoursPerWeek# * 52
        End If
    End If
End If

' dkostka - 02/18/2002 - Added Val(Format(x, "@")) around expression to replace null with 0.
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 11 Then
'If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 15 Then
If dynSH_Job1("JB_MIDPOINT") >= 1 And dynSH_Job1("JB_MIDPOINT") <= 20 Then
    Jb_No = Val(Format(dynSH_Job1("JB_S" & dynSH_Job1("JB_MIDPOINT")), "@"))
End If

fglbCOMPA# = 0

If Jb_No <> 0 And dblSsalary# <> 0 Then 'laura 03/23/98
  fglbCOMPA# = (dblSsalary# / Jb_No) * 100
End If

 
If fglbCOMPA# > 999.99 Then fglbCOMPA# = 999.99


fglbGRADE$ = "00"
xSalGrade = dblNewSalary
'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
'For X% = 1 To 11
'For X% = 1 To 15
For x% = 1 To 20
    'Added D by Bryan 28/Sep/05 Ticket#9354
    If IsNumeric(dynSH_Job1("JB_S" & Format(x%, "##"))) Then
        If snapJob("JB_SALCD") = "H" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                xSalGrade = snapJob("JB_S" & Format(x%, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                xSalGrade = (snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)) / 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                xSalGrade = snapJob("JB_S" & Format(x%, "##")) * (dblHoursPerWeek# * 52)
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 366 / (dblHoursPerWeek# * 52)
                Else
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 365 / (dblHoursPerWeek# * 52)
                End If
                
                'Ticket #17654 - formula correction
                xSalGrade = snapJob("JB_S" & Format(x%, "##")) * fglbDhrs
            End If
        ElseIf snapJob("JB_SALCD") = "A" Then
            If dynSH_Job1("SH_SALCD") = "H" Then
                If dblHoursPerWeek# = 0 Then
                    xSalGrade = 0
                Else
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52)
                End If
            ElseIf dynSH_Job1("SH_SALCD") = "M" Then
                xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 12
            ElseIf dynSH_Job1("SH_SALCD") = "A" Then
                xSalGrade = snapJob("JB_S" & Format(x%, "##"))
            ElseIf dynSH_Job1("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 366
                Else
                    xSalGrade = snapJob("JB_S" & Format(x%, "##")) * 365
                End If
                
                'Ticket #17654 - formula correction
                xSalGrade = snapJob("JB_S" & Format(x%, "##")) / (dblHoursPerWeek# * 52) * fglbDhrs
            End If
        End If
        If dblNewSalary >= xSalGrade And dynSH_Job1("JB_S" & Format(x%, "##")) > 0 Then
            cX$ = CStr(x)
            If x% <= 9 Then cX$ = "0" & cX$
            fglbGRADE$ = cX$
        End If
    End If
Next x%

If IsNumeric(dynSH_Job1("JB_S1")) Then
    If dblSsalary# < dynSH_Job1("JB_S1") Then
        fglbGRADE$ = "00"
    End If
End If
snapJob.Close
End Sub

Private Sub Upd_Salary_Data() 'Frank 4/12/2000
Dim fTablSalHis As New ADODB.Recordset ', dynSH_Job1 As Recordset
Dim SQLQ, lngLastCurrentID&
Dim dblOSalary, dblNewSalary
'Dim EmpNo&
Dim prec%, pct%
Dim xSHID 'George added on MAr 10,2006 #9965
Dim xStr

MailBody = ""
If glbOracle Then
    If glbMultiGrid Then
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY,HR_JOB_HISTORY,HRJOB_GRADE"
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HR_JOB_HISTORY.JH_GRID "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HRJOB_GRADE.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HRJOB_GRADE.JB_GRID "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    Else
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB.* "
        SQLQ = SQLQ & " FROM HR_SALARY_HISTORY,HR_JOB_HISTORY,HRJOB"
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    End If
Else
    If glbMultiGrid Then
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB_GRADE.* "
        SQLQ = SQLQ & " FROM (HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HR_JOB_HISTORY.JH_GRID "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) "
        SQLQ = SQLQ & " INNER JOIN HRJOB_GRADE"
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HRJOB_GRADE.JB_CODE "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_GRID = HRJOB_GRADE.JB_GRID "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    Else
        SQLQ = "SELECT HR_SALARY_HISTORY.*, HRJOB.* "
        SQLQ = SQLQ & " FROM (HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HR_JOB_HISTORY.JH_JOB AND HR_SALARY_HISTORY.SH_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR"
        If glbMulti Then
            SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_SDATE=HR_JOB_HISTORY.JH_SDATE"
        End If
        SQLQ = SQLQ & " ) "
        SQLQ = SQLQ & " INNER JOIN HRJOB"
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & " WHERE HR_SALARY_HISTORY.SH_CURRENT <>0 AND HR_JOB_HISTORY.JH_CURRENT<>0"
    End If
End If
'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
'update on the salary screen. So changed from < to <=.
SQLQ = SQLQ & " AND SH_EDATE <= " & Date_SQL(dlpEDate.Text)
SQLQ = SQLQ & " AND " & WSQLQ
SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn

'Ticket #28969 - Employment Status
If Len(clpCode(0).Text) > 0 Then
    SQLQ = SQLQ & " AND ED_EMP IN ('" & Replace(clpCode(0).Text, ",", "','") & "') "
End If
SQLQ = SQLQ & ")"

If glbNoNONE And glbNoEXEC Then 'Hemu -EXE
    SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG <> 'NONE' AND ED_ORG <> 'EXEC') "
ElseIf glbNoNONE Then
    SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG <> 'NONE') "
ElseIf glbNoEXEC Then    'Hemu -EXE
    SQLQ = SQLQ & " AND SH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG <> 'EXEC') " 'Hemu -EXE
End If
If dynSH_Job1.State <> 0 Then dynSH_Job1.Close
dynSH_Job1.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
lngRecs& = dynSH_Job1.RecordCount

'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
'update on the salary screen.
'MsgSal = "The effective date is same or later than" & Chr(10) & "your most current record "
MsgSal = "The effective date is earlier than" & Chr(10) & "your most current record."
MsgSal = MsgSal & Chr(10) & "Employee's Number:"
IfDisplay = False

Do Until dynSH_Job1.EOF
    prec% = prec% + 1
    pct% = Int(100 * (prec% / (lngRecs&)))
    MDIMain.panHelp(0).FloodPercent = pct%
    EmpNo& = dynSH_Job1("SH_EMPNBR")
    
    'Ticket #18668 and Ticket #19154 - Allow same salary effective date update since we are allowing manual
    'update on the salary screen. So changed from >= to >.
    If dynSH_Job1("SH_EDATE") > CVDate(dlpEDate.Text) Then
        MsgSal = MsgSal & Chr(10) & "                        " & Str(EmpNo&) '& Chr(10)
        IfDisplay = True
        GoTo NextRec
    End If
    lngLastCurrentID& = dynSH_Job1("SH_ID")
    
    If fTablSalHis.State <> 0 Then fTablSalHis.Close
    fTablSalHis.Open "SELECT * FROM HR_SALARY_HISTORY WHERE SH_ID = " & lngLastCurrentID&, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    OSalary = fTablSalHis("SH_SALARY")
    OTOTAL = fTablSalHis("SH_TOTAL")
    oPayP = fTablSalHis("SH_PAYP")
    OEDate = fTablSalHis("SH_EDATE")
    ONDate = fTablSalHis("SH_NEXTDAT")
    OJOB1 = fTablSalHis("SH_JOB")
    oGrid = fTablSalHis("SH_GRID")
    oPayrollID = fTablSalHis("SH_PAYROLL_ID")
    OSalCD = fTablSalHis("SH_SALCD")
    oStep = fTablSalHis("SH_GRADE")
    
    If Len(fTablSalHis("SH_WHRS")) < 1 Then
        dblWHours# = 0
    Else
        dblWHours# = fTablSalHis("SH_WHRS")
    End If
    fTablSalHis("SH_CURRENT") = False
    fTablSalHis.Update
    
    'Ticket #16991 - Do not update Vadim's HR_EMP_HISTORY table because the Rate level of the employee is
    'remaining same, only the actual salary is changing and this table only stores the Rate Level
    'Comment enhacement - Ticket #16115
    'City of Niagara Falls - Ticket #15542
    'If glbVadim And glbCompSerial = "S/N - 2276W" Then
    '    'Update previous salary record in Vadim's HR_EMP_HIST table with End Date
    '    Call Update_VadimDB_HR_EMP_HISTORY(oPayrollID, OEDate, "", "", "", "M", DateAdd("d", -1, CVDate(dlpEDate)))
    'End If
    
    fTablSalHis.AddNew
    fTablSalHis("SH_COMPNO") = dynSH_Job1("SH_COMPNO")
    fTablSalHis("SH_EMPNBR") = dynSH_Job1("SH_EMPNBR")
    fTablSalHis("SH_EDATE") = CVDate(dlpEDate)
    'Add by Frank Jan 10,2002 As Jerry request
    If IsDate(ONDate) Then
        If CVDate(ONDate) > CVDate(dlpEDate) Then
            UpdateFollowup dynSH_Job1("SH_EMPNBR"), CVDate(ONDate), CVDate(ONDate), "SREV"
            fTablSalHis("SH_NEXTDAT") = ONDate
        End If
    End If
    'Add by Frank Jan 10,2002 As Jerry request
    fTablSalHis("SH_CURRENT") = True
    fTablSalHis("SH_SDATE") = dynSH_Job1("SH_SDATE")
    fTablSalHis("SH_SALCD") = dynSH_Job1("SH_SALCD")
    fTablSalHis("SH_WHRS") = dynSH_Job1("SH_WHRS")
    fTablSalHis("SH_PAYP") = dynSH_Job1("SH_PAYP")
    fTablSalHis("SH_PAYP_TABLE") = dynSH_Job1("SH_PAYP_TABLE")
    fTablSalHis("SH_SREAS_TABLE") = dynSH_Job1("SH_SREAS_TABLE")
    dblOSalary = dynSH_Job1("SH_SALARY")
    If optDollars Then
        dblNewSalary = dblOSalary + medChng
        If optStep Then
            dblNewSalary = GetNewStepSalary_1(Round2DEC(dblNewSalary, dynSH_Job1("SH_EMPNBR")))
        End If
    End If
    If optPct Then
        dblNewSalary = dblOSalary * ((medChng / 100) + 1)
        If optStep Then
            dblNewSalary = GetNewStepSalary_1(Round2DEC(dblNewSalary, dynSH_Job1("SH_EMPNBR")))
        End If
    End If
    fTablSalHis("SH_SALARY") = Round2DEC(dblNewSalary, dynSH_Job1("SH_EMPNBR"))
    fTablSalHis("SH_JOB") = dynSH_Job1("SH_JOB")
    fTablSalHis("SH_GRID") = dynSH_Job1("SH_GRID")
    fTablSalHis("SH_PAYROLL_ID") = dynSH_Job1("SH_PAYROLL_ID")
    fTablSalHis("SH_JOB_ID") = dynSH_Job1("SH_JOB_ID")

    Call modSetCOMPA_GRADE_1(Round2DEC(dblNewSalary, dynSH_Job1("SH_EMPNBR"))) ' sets fglbCOMPA#, and fglbGRADE
    fTablSalHis("SH_COMPA") = Round(fglbCOMPA#, 2)
    fTablSalHis("SH_GRADE") = Format(fglbGRADE$, "00")

    fTablSalHis("SH_SREAS1") = clpCode(4).Text
    If dblOSalary <> 0 Then fTablSalHis("SH_SALPC1") = (dblNewSalary - dblOSalary) / dblOSalary
    fTablSalHis("SH_SALCHG1") = dblNewSalary - dblOSalary

    fTablSalHis("SH_TRANSDATE") = Now
    fTablSalHis("SH_LDATE") = Now
    fTablSalHis("SH_LTIME") = Time$
    fTablSalHis("SH_LUSER") = glbUserID 'glbLEE_ID
    If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
        fTablSalHis("SH_PREMIUM") = dynSH_Job1("SH_PREMIUM")
        fTablSalHis("SH_TOTAL") = fTablSalHis("SH_SALARY") + dynSH_Job1("SH_PREMIUM") 'dynSH_Job1("SH_TOTAL")
        fTablSalHis("SH_VGROUP") = dynSH_Job1("SH_VGROUP")
        fTablSalHis("SH_VSTEP") = dynSH_Job1("SH_VSTEP")
    End If
    fTablSalHis.Update
    
    If gsEMAIL_ONSALARY Then
        MailBody = MailBody & GetEmpName(dynSH_Job1("SH_EMPNBR")) & vbCrLf
    End If
    
    xSHID = fTablSalHis("SH_ID") 'George added on MAr 10,2006 #9965
    
    If glbVadim Then Call Transfer_Salary(fTablSalHis)
    
    'Ticket #28595 - Update employee's Attendance records as welll if selected
    If chkUpdAttendance Then Call Update_Attendance_SalaryInfo(fTablSalHis)
    
    Call updBenefitForSalDEPN(EmpNo&)   'jaddy 9/10/99
    
    'Ticket #16991 - Do not update Vadim's HR_EMP_HISTORY table because the Rate level of the employee is
    'remaining same, only the actual salary is changing and this table only stores the Rate Level
    'City of Niagara Falls - Ticket #15542
    'If glbVadim And glbCompSerial = "S/N - 2276W" Then
    '    'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
    '    Call Update_VadimDB_HR_EMP_HISTORY(fTablSalHis("SH_PAYROLL_ID"), CVDate(dlpEDate), "", Val(fglbGRADE$), fTablSalHis("SH_JOB"), "A")
    'End If
    
    Call Employee_Master_Integration(EmpNo&)
    If glbGP Then
        Call Salary_Integration(EmpNo&, , False, True, xSHID) 'George added on MAr 10,2006 #9965
    Else
        Call Salary_Integration(EmpNo&) 'Ticket #15646
    End If
    NSalary = dblNewSalary
    NEDate = CVDate(dlpEDate.Text)
    NNDate = ""
    If Not AUDITSALY() Then MsgBox "ERROR - AUDIT FILE"
NextRec:
    dynSH_Job1.MoveNext
Loop

If IfDisplay = True Then
    MsgBox MsgSal, vbInformation, "Update Salary"
End If

If gsEMAIL_ONSALARY Then
     If Len(MailBody) > 0 Then
        If prec% = 1 Then
            xStr = "The following employee's salary has "
        Else
            xStr = "The following employee's salaries have "
        End If
        If optDollars Then
            xStr = xStr & " been increased by $" & medChng & "." & vbCrLf '& vbCrLf
        End If
        If optPct Then
            xStr = xStr & " been increased by " & medChng & "%." & vbCrLf  '& vbCrLf
        End If
        xStr = xStr & "Reason: " & GetTABLDesc("SDRC", clpCode(4)) & vbCrLf
        xStr = xStr & "Effective Date: " & dlpEDate & vbCrLf & vbCrLf
        MailBody = xStr & MailBody
        Screen.MousePointer = DEFAULT
        Call imgEmail_Click
     End If
     
End If

If prec% > 0 Then
    If prec% = 1 Then
        MsgBox prec% & " employee salary record was updated"
    Else
        MsgBox prec% & " employees salary records were updated"
    End If
End If
End Sub

Private Function GetEmpName(xEmpNo)
Dim rsTemp As New ADODB.Recordset
Dim xStr, SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & xEmpNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xStr = "Employee #:" & (xEmpNo) & " Name: " & rsTemp("ED_FNAME") & " " & rsTemp("ED_SURNAME")
    End If
    rsTemp.Close
    GetEmpName = xStr
End Function

Public Sub imgEmail_Click()
Dim xEmail
On Error GoTo Email_Err
    If gsEMAIL_ONSALARY Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        xEmail = GetComPreferEmail("EMAIL_ONSALARY")
        
        If Len(xEmail) > 0 Then
            frmSendEmail.txtTo.Text = xEmail 'GetComPreferEmail("EMAIL_ONSALARY")
            'frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            frmSendEmail.txtSubject.Text = "info:HR Salary Change Notice"
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

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUJOBS"
End Sub

Private Sub Form_Load()

glbOnTop = "FRMUJOBS"

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Screen.MousePointer = HOURGLASS

'Release 8.0 - Since multiple select, this is not required
'If glbCompSerial = "S/N - 2191W" Then clpCode(1).MaxLength = 6

cmbRound.AddItem "Yes"
cmbRound.AddItem "No"
cmbRound.ListIndex = 1

cmbPrecision.AddItem "0"
cmbPrecision.AddItem "2"
cmbPrecision.AddItem "3"
cmbPrecision.AddItem "4"

Call setRptCaption(Me)

If glbMultiGrid Then
    lblGrid.Visible = True
    clpGrid.Visible = True
End If
If glbSyndesis Then
    lblPosGroup.Caption = "Position Grade"
    clpCode(1).Tag = "00-Enter Position Grade"
End If

If glbCompSerial = "S/N - 2259W" Then '#15908
    lblPosStatus.Caption = "Income Code"
    clpCode(3).TABLTitle = "Income Code"
End If

If glbCompSerial = "S/N - 2172W" Then 'Lanark Ticket #17221 by Frank 08/19/2009
    lblPosGroup.Caption = "Salary Level"
    clpCode(1).TABLTitle = "Salary Level Code"
End If

If glbCompDecHR = 3 Then medChng.Format = "#,##0.000;(#,##0.000)"
If glbCompDecHR = 4 Then medChng.Format = "#,##0.0000;(#,##0.0000)"
If glbCompDecHR = 2 Then medChng.Format = "Fixed"
Call INI_Controls(Me)

If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    clpJob.TransDiv = glbWFCUserSecList
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
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
Set frmUJobs = Nothing
End Sub


Private Sub medChng_GotFocus()
Call SetPanHelp(ActiveControl)
If optPct Then
    medChng.Format = "0.00" 'changed by RAUBREY 6/3/97
    
Else
    If glbCompDecHR = 3 Then medChng.Format = "#,##0.000;(#,##0.000)"
    If glbCompDecHR = 4 Then medChng.Format = "#,##0.0000;(#,##0.0000)"
    If glbCompDecHR = 2 Then medChng.Format = "Fixed"
End If

End Sub



Private Sub medChng_LostFocus()
If Not optPct Then
    If glbCompDecHR = 3 Then medChng.Format = "#,##0.000;(#,##0.000)"
    If glbCompDecHR = 4 Then medChng.Format = "#,##0.0000;(#,##0.0000)"
    If glbCompDecHR = 2 Then medChng.Format = "Fixed"
Else
   ' medchng.Format = "Percent"'changed by RAUBREY 6/3/97
   medChng.Format = "0.00" 'changed by RAUBREY 6/3/97
   
End If

End Sub

Private Sub modJobUpd()
Dim Msg$, DgDef As Variant, Response%, noRecs&
Dim dyn_HRJOB As New ADODB.Recordset
Dim dyn_HRJOB_GRADE As New ADODB.Recordset
Dim SQLQ, x%, strFld
Dim Y, xCount
Dim RoundSal As Double, strRoundSal As String, strFirst As String   'laura 03/26/98


On Error GoTo cmdUpdErr

Screen.MousePointer = HOURGLASS

xCount = 0
WSQLQ = "1=1 "

If glbMultiGrid Then
    If clpCode(1).Text <> "" Or clpCode(3).Text <> "" Then
        WSQLQ = WSQLQ & " AND JB_CODE IN (SELECT JB_CODE FROM HRJOB WHERE 1=1 "
        If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND JB_GRPCD IN ('" & Replace(clpCode(1).Text, ",", "','") & "') "
        If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND JB_STATUS IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
        WSQLQ = WSQLQ & ")"
    End If
Else
    If clpCode(1).Text <> "" Then WSQLQ = WSQLQ & " AND JB_GRPCD IN ('" & Replace(clpCode(1).Text, ",", "','") & "') "
    If clpCode(3).Text <> "" Then WSQLQ = WSQLQ & " AND JB_STATUS IN ('" & Replace(clpCode(3).Text, ",", "','") & "') "
End If
If clpJob.Text <> "" Then WSQLQ = WSQLQ & " AND JB_CODE IN ('" & Replace(clpJob.Text, ",", "','") & "') "
If clpCode(2).Text <> "" Then WSQLQ = WSQLQ & " AND JB_ORG IN ('" & Replace(clpCode(2).Text, ",", "','") & "') "
If clpGrid.Text <> "" Then WSQLQ = WSQLQ & " AND JB_GRID IN ('" & Replace(clpGrid.Text, ",", "','") & "') "

If glbMultiGrid Then
    SQLQ = "SELECT * FROM HRJOB_GRADE WHERE " & WSQLQ
Else
    SQLQ = "SELECT * FROM HRJOB WHERE " & WSQLQ
End If
dyn_HRJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If Not dyn_HRJOB.EOF And Not dyn_HRJOB.BOF Then
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(0).FloodPercent = 0
    noRecs& = dyn_HRJOB.RecordCount
    Screen.MousePointer = DEFAULT
    Msg$ = "Are you sure you wish to update the " & CStr(noRecs&)
    Msg$ = Msg$ & Chr(10) & "Jobs within this selection by "
    If optPct Then
         Msg$ = Msg$ & Chr(10) & CStr(medChng) & " Percent?"
    Else
        Msg$ = Msg$ & Chr(10) & CStr(medChng) & " Dollars?"
    End If
    

    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
    Response% = MsgBox(Msg$, DgDef, "Warning!")
    If Response% = IDNO Then
        Exit Sub
    End If
    Screen.MousePointer = HOURGLASS
    'BeginTrans 'Ticket #15590 Commented by Frank,
                'Great Plains integration needs update the Position Master for each code during this loop
    While Not dyn_HRJOB.EOF
        'dyn_HRJOB.Edit
        If optPct Then
            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'For X% = 1 To 11
            'For X% = 1 To 15
            For x% = 1 To 20
                strFld = "JB_S" & CStr(x%)
                If Not IsNull(dyn_HRJOB(strFld)) Then 'changed by RAUBREY 6/3/97
                  If Len(dyn_HRJOB(strFld)) > 0 Then
                    If dyn_HRJOB(strFld) > 0 Then
                        Y = (medChng / 100) + 1  'changed by RAUBREY 6/3/97
                        '~~~~~~~ ADDED BY RAUBREY 8/19/97 ~~~~~~~~~~~~~~
                        Y = dyn_HRJOB(strFld) * Y
                        Y = Round2DEC(Y)
                        'laura 03/26/98
                        'If glbCompSerial = "S/N - 2191W" Then
                          If cmbRound.ListIndex = 0 Then
                              'Y = CLng(Y)  'Ticket #14699
                              Y = Round(Y, cmbPrecision.Text)  'Ticket #14699
                          Else
                              Y = Y
                          End If
                        'End If
                        '~~~~~~~~~
                        If glbVadim And Not glbMultiGrid Then
                            'Hemu - Town of Aurora only - Ticket #10263
                            If glbCompSerial = "S/N - 2378W" Then
                                If dyn_HRJOB("JB_SALCD") = "A" Then
                                    If IsNumeric(dyn_HRJOB("JB_FTEHRS")) And dyn_HRJOB("JB_FTEHRS") <> 0 Then
                                        Call Passing_Salary_Grid_Vadim(x%, dyn_HRJOB(strFld & "A").Value, Y / dyn_HRJOB("JB_FTEHRS"), Date, dyn_HRJOB("JB_CODE").Value)
                                    End If
                                Else
                                    Call Passing_Salary_Grid_Vadim(x%, dyn_HRJOB(strFld).Value, Y, Date, dyn_HRJOB("JB_CODE").Value)
                                End If
                            Else
                                Call Passing_Salary_Grid_Vadim(x%, dyn_HRJOB(strFld).Value, Y, Date, dyn_HRJOB("JB_CODE").Value)
                            End If
                        End If
                        DoEvents
                        dyn_HRJOB(strFld) = Y
                        If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                            If IsNumeric(dyn_HRJOB("JB_FTEHRS")) And dyn_HRJOB("JB_FTEHRS") <> 0 Then
                                If dyn_HRJOB("JB_SALCD") = "A" Then
                                    dyn_HRJOB(strFld & "A") = Y / dyn_HRJOB("JB_FTEHRS")  'To get Hourly Rate
                                ElseIf dyn_HRJOB("JB_SALCD") = "H" Then
                                    dyn_HRJOB(strFld & "A") = Y * dyn_HRJOB("JB_FTEHRS")  'To get Annual Amount
                                End If
                            End If
                        End If
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                      ' MsgBox Str(dyn_HRJOB(strFld))
                    
                    End If
                  End If
                End If
            Next x%
        Else
            'Ticket #22682 - Release 8.0: Increased Grid Steps to 15 -> 20
            'For X% = 1 To 11
            'For X% = 1 To 15
            For x% = 1 To 20
                strFld = "JB_S" & CStr(x%)
                If Not IsNull(dyn_HRJOB(strFld)) Then 'changed by RAUBREY 6/3/97
                  If Len(dyn_HRJOB(strFld)) > 0 Then
                    If dyn_HRJOB(strFld) > 0 Then 'changed by RAUBREY 6/3/97
                        '~~~~~~~ ADDED BY RAUBREY 8/19/97 ~~~~~~~~~~~~~~
                        Y = dyn_HRJOB(strFld) + medChng
                        Y = Round2DEC(Y)
                        'laura 03/26/98
                        'If glbCompSerial = "S/N - 2191W" Then
                          If cmbRound.ListIndex = 0 Then
                              'Y = CLng(Y)   'Ticket #14699
                              Y = Round(Y, cmbPrecision.Text)  'Ticket #14699
                          Else
                              Y = Y
                          End If
                        'End If

                        '~~~~~~~~~
                        If glbVadim And Not glbMultiGrid Then
                            'Hemu - Town of Aurora only - Ticket #10263
                            If glbCompSerial = "S/N - 2378W" Then
                                If dyn_HRJOB("JB_SALCD") = "A" Then
                                    If IsNumeric(dyn_HRJOB("JB_FTEHRS")) And dyn_HRJOB("JB_FTEHRS") <> 0 Then
                                        Call Passing_Salary_Grid_Vadim(x%, dyn_HRJOB(strFld & "A"), Y / dyn_HRJOB("JB_FTEHRS"), Date, dyn_HRJOB("JB_CODE").Value)
                                    End If
                                Else
                                    Call Passing_Salary_Grid_Vadim(x%, dyn_HRJOB(strFld).Value, Y, Date, dyn_HRJOB("JB_CODE").Value)
                                End If
                            Else
                                Call Passing_Salary_Grid_Vadim(x%, dyn_HRJOB(strFld).Value, Y, Date, dyn_HRJOB("JB_CODE").Value)
                            End If
                        End If
                        DoEvents
                        dyn_HRJOB(strFld) = Y
                        If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora
                            If IsNumeric(dyn_HRJOB("JB_FTEHRS")) And dyn_HRJOB("JB_FTEHRS") <> 0 Then
                                If dyn_HRJOB("JB_SALCD") = "A" Then
                                    dyn_HRJOB(strFld & "A") = Y / dyn_HRJOB("JB_FTEHRS")  'To get Hourly Rate
                                ElseIf dyn_HRJOB("JB_SALCD") = "H" Then
                                    dyn_HRJOB(strFld & "A") = Y * dyn_HRJOB("JB_FTEHRS")  'To get Annual Amount
                                End If
                            End If
                        End If
                        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                    End If
                  End If
                End If
            Next x%
        End If

        dyn_HRJOB("JB_LDATE") = Now
        dyn_HRJOB("JB_LTIME") = Time$
        dyn_HRJOB("JB_LUSER") = glbUserID
        dyn_HRJOB.Update
        
        If glbGP Then 'Ticket #15590
            Call Codes_Master_Integration("POSITION", dyn_HRJOB("JB_CODE"))
        End If
        dyn_HRJOB.MoveNext
    Wend
    'CommitTrans
    If chkSalary Then Call Upd_Salary_Data 'FRANK 4/12/2000
    If chkCompa Then Call Upd_Related_Data
    'gdbAdoIhr001.Execute "UPDATE HRJOB SET JB_LUSER=" & glbLEE_ID & " WHERE JB_LUSER=999999998"
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).FloodShowPct = False
    MsgBox "Update completed"
   
Else
    MsgBox lStr("There were no Positions within that Group")
End If

Screen.MousePointer = DEFAULT

Exit Sub
cmdUpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Error", "HRJobs", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub optAmount_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optDollars_Click(Value As Integer)
If Value Then
    If glbCompDecHR = 3 Then medChng.Format = "#,##0.000;(#,##0.000)"
    If glbCompDecHR = 4 Then medChng.Format = "#,##0.0000;(#,##0.0000)"
    If glbCompDecHR = 2 Then medChng.Format = "Fixed"
Else
  ' medchng.Format = "Percent"'changed by RAUBREY 6/3/97
  medChng.Format = "0.00" 'changed by RAUBREY 6/3/97
  
End If

End Sub

Private Sub optDollars_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optDollars_LostFocus()
If optDollars Then
    If glbCompDecHR = 3 Then medChng.Format = "#,##0.000;(#,##0.000)"
    If glbCompDecHR = 4 Then medChng.Format = "#,##0.0000;(#,##0.0000)"
    If glbCompDecHR = 2 Then medChng.Format = "Fixed"
Else
  ' medchng.Format = "Percent"'changed by RAUBREY 6/3/97
  medChng.Format = "0.00" 'changed by RAUBREY 6/3/97
  
End If

End Sub

Private Sub optPct_Click(Value As Integer)
If Not Value Then
    If glbCompDecHR = 3 Then medChng.Format = "#,##0.000;(#,##0.000)"
    If glbCompDecHR = 4 Then medChng.Format = "#,##0.0000;(#,##0.0000)"
    If glbCompDecHR = 2 Then medChng.Format = "Fixed"
Else
    'medchng.Format = "Percent"'changed by RAUBREY 6/3/97
    medChng.Format = "0.00" 'changed by RAUBREY 6/3/97
End If

End Sub

Private Sub optPct_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optPct_LostFocus()
If optDollars Then
    If glbCompDecHR = 3 Then medChng.Format = "#,##0.000;(#,##0.000)"
    If glbCompDecHR = 4 Then medChng.Format = "#,##0.0000;(#,##0.0000)"
    If glbCompDecHR = 2 Then medChng.Format = "Fixed"
Else
  ' medchng.Format = "Percent"'changed by RAUBREY 6/3/97
   medChng.Format = "0.00" 'changed by RAUBREY 6/3/97
   
End If

End Sub


Private Function Round2DEC(tmpNUM, Optional xEmpNo) 'laura nov 10, 1997
Dim strNUM As String, x%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" And Not (IsMissing(xEmpNo)) Then 'City of Timmins
    If GetEmpData(xEmpNo, "ED_REGION") <> "S" Then
        Round2DEC = Round(tmpNUM, 2)
    Else
        Round2DEC = Round(tmpNUM, glbCompDecHR)
    End If
Else
    Round2DEC = Round(tmpNUM, glbCompDecHR)
End If

End Function


Private Sub panUpd_Click()

End Sub

Private Sub optStep_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Public Sub UpdateFollowup(EmpNbr, OldDate, NewDate, Code)
    Dim rsFollow As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR=" & EmpNbr
    If Not IsNull(OldDate) Then
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(OldDate)
    End If
    
    rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsFollow.EOF Or rsFollow.BOF Or IsNull(OldDate) Then
        rsFollow.AddNew
        rsFollow("EF_EMPNBR") = EmpNbr
        rsFollow("EF_FREAS") = Code
        rsFollow("EF_COMPLETED") = False
    End If
    rsFollow("EF_FDATE") = NewDate
    rsFollow("EF_LDATE") = Date
    rsFollow("EF_LTIME") = Format(Now, "Medium Time")
    rsFollow.Update
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
UpdateRight = True
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Private Sub Transfer_Salary(rsNew As ADODB.Recordset)
    Dim rsEmpJob As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim HRChanges As New Collection
    Dim UptSalaryDate As Date
    Dim HRSalary As New Collection
    Dim xEmpNbr
    Dim xPayrollID
    Dim xPHrs
    Dim xWHrs, xNiagaraWHRS
    Dim xEDate
    Dim xSalCd
    Dim UpdateAudit
    
    xEmpNbr = rsNew("SH_EMPNBR")
    If rsNew("SH_PAYROLL_ID") = "" Or IsNull(rsNew("SH_PAYROLL_ID")) Then
        xPayrollID = GetEmpData(rsNew("SH_EMPNBR"), "ED_PAYROLL_ID")
    Else
        xPayrollID = rsNew("SH_PAYROLL_ID")
    End If
    xEDate = rsNew("SH_EDATE")
    
    rsEmpJob.Open "SELECT JH_ID,JH_JOB,JH_DHRS,JH_PHRS,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEmpNbr & " AND JH_PAYROLL_ID='" & xPayrollID & "'", gdbAdoIhr001, adOpenForwardOnly
    xPHrs = 0
    xWHrs = 0
    If Not rsEmpJob.EOF Then
        xPHrs = Val(rsEmpJob("JH_PHRS") & "")
        xWHrs = Val(rsEmpJob("JH_WHRS") & "") 'Hemu - it was asssigning JH_DHRS - it should pass Weekly Hours
        xNiagaraWHRS = Val(rsEmpJob("JH_WHRS") & "")
        
        'City of Niagara Falls  = Dhrs = Hours Per Days from Position Master, fglbNiagPhrs = Pay Period
        If glbCompSerial = "S/N - 2276W" Then
            rsSal.Open "SELECT SH_EMPNBR, SH_PAYP, SH_WHRS FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & xEmpNbr & " AND SH_PAYROLL_ID = '" & xPayrollID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                xPHrs = Val(rsSal("SH_PAYP") & "")
                xNiagaraWHRS = Val(rsSal("SH_WHRS") & "")
            End If
            rsSal.Close
            Set rsSal = Nothing
            xWHrs = GetJobData(rsEmpJob("JH_JOB"), "JB_DHRS", 1)
            xWHrs = Val(xWHrs & "")
        End If
    End If
    rsEmpJob.Close
   
    If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
        If isChanged_Salary(HRSalary, OTOTAL, rsNew("SH_TOTAL"), True) Then UpdateAudit = True
    Else
        If isChanged_Salary(HRSalary, OSalary, rsNew("SH_SALARY"), True) Then UpdateAudit = True
    End If
    If isChanged_Salary(HRSalary, OSalCD, rsNew("SH_SALCD")) Then UpdateAudit = True
    
    If glbVadim And UpdateAudit Then
        'Ticket #21352 - City of Kawartha Lakes
        If glbCompSerial = "S/N - 2363W" Then
            Call Passing_Salary_Vadim(HRSalary, Salary, Date, xPHrs, xWHrs, xEmpNbr, xPayrollID, , xNiagaraWHRS)
        Else
            Call Passing_Salary_Vadim(HRSalary, Salary, xEDate, xPHrs, xWHrs, xEmpNbr, xPayrollID, , xNiagaraWHRS)
        End If
    End If
    
    'Ticket #24565 - District Municipality of Muskoka
    If glbCompSerial = "S/N - 2373W" Then
        'They want to transfer for 181W as well now - Nov 3rd 2014
        'Ticket #24565 - if Union = '181W' then do not transfer Probation Date, Level and After Probation
        'If GetEmpData(xEmpnbr, "ED_ORG") = "181W" Then
        '    'Do not transfer Probation Date, Level and After Probation
        'Else
            If isChanged_Field(HRChanges, oStep, rsNew("SH_GRADE"), True) Then Debug.Print "" ' do nothing for the audit transfer
        'End If
    Else
        'Ticket #25469 - City of Campbell River - do not transfer Probation levels
        If glbCompSerial <> "S/N - 2458W" Then
            If isChanged_Field(HRChanges, oStep, rsNew("SH_GRADE"), True) Then Debug.Print "" ' do nothing for the audit transfer
        End If
    End If
    
    If isChanged_Field(HRChanges, OEDate, rsNew("SH_EDATE")) Then UpdateAudit = True
    If glbCompSerial <> "S/N - 2373W" Then 'DMuskoka - Ticket #24565 - Do not transfer Next Review Date
        If isChanged_Field(HRChanges, ONDate, rsNew("SH_NEXTDAT")) Then UpdateAudit = True
    End If
    Call Passing_Changes(HRChanges, Salary, "M", Date, xEmpNbr, xPayrollID)

End Sub
