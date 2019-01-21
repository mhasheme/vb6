VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRPOE 
   Caption         =   "Plan of Establishment"
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2160
      TabIndex        =   7
      Tag             =   "00-Enter Employee Number"
      Top             =   2760
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Tag             =   "00-Enter Position Code "
      Top             =   3085
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   3
      Tag             =   "00-Enter Location Code"
      Top             =   1455
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpGLNum 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "00-Specific GL No Desired"
      Top             =   1130
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin VB.Frame frmLegend 
      Caption         =   "Legend"
      Height          =   1335
      Left            =   360
      TabIndex        =   31
      Top             =   4800
      Width           =   8655
      Begin VB.Label lblLegend4 
         Caption         =   "If Selection Criteria entered is using data in the Purple or Black, at least one Red Criteria must be entered"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   960
         Width           =   8055
      End
      Begin VB.Label lblLegend3 
         Caption         =   "Black - Employee Master"
         Height          =   255
         Left            =   720
         TabIndex        =   34
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblLegend2 
         Caption         =   "Purple - Position History"
         Height          =   255
         Left            =   720
         TabIndex        =   33
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label lblLegend1 
         Caption         =   "Red - Budgeted Position Master"
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   240
         Width           =   2775
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   16
      Top             =   6585
      Width           =   10785
      _Version        =   65536
      _ExtentX        =   19024
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
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
         Left            =   2085
         TabIndex        =   15
         Tag             =   "Print the report"
         Top             =   150
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdView 
         Appearance      =   0  'Flat
         Caption         =   "&View"
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
         Left            =   1080
         TabIndex        =   13
         Tag             =   "View this report (with option to print)"
         Top             =   150
         Visible         =   0   'False
         Width           =   855
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
         Left            =   120
         TabIndex        =   14
         Tag             =   "Close and exit this screen"
         Top             =   150
         Visible         =   0   'False
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   3480
         Top             =   150
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
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   2160
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
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2160
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   4
      Tag             =   "00-Enter Union Code"
      Top             =   1785
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
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Tag             =   "00-Enter Employee Status Code"
      Top             =   2115
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
      Index           =   3
      Left            =   2160
      TabIndex        =   9
      Tag             =   "00-Enter Position Group"
      Top             =   3410
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   10
      Tag             =   "00-Enter Region Code"
      Top             =   3730
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   2160
      TabIndex        =   11
      Tag             =   "00-Enter Administered By Code"
      Top             =   4055
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   2160
      TabIndex        =   12
      Tag             =   "00-Enter Section Code"
      Top             =   4360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Tag             =   "EDPT-Category"
      Top             =   2435
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.Label lblGL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "G/L Code"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   420
      TabIndex        =   30
      Top             =   1130
      Width           =   870
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   420
      TabIndex        =   29
      Top             =   480
      Width           =   555
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   420
      TabIndex        =   28
      Top             =   810
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   27
      Top             =   1790
      Width           =   420
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   26
      Top             =   2120
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
      Left            =   420
      TabIndex        =   25
      Top             =   2760
      Width           =   1290
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   24
      Top             =   3410
      Width           =   1035
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
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   420
      TabIndex        =   22
      Top             =   1460
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
      Left            =   420
      TabIndex        =   21
      Top             =   3730
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
      Left            =   420
      TabIndex        =   20
      Top             =   4055
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
      Left            =   420
      TabIndex        =   19
      Top             =   4360
      Width           =   540
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   18
      Top             =   2435
      Width           =   630
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   420
      TabIndex        =   17
      Top             =   3085
      Width           =   975
   End
End
Attribute VB_Name = "frmRPOE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Job_Snap As New ADODB.Recordset
Dim LGR_snap As New ADODB.Recordset
Dim CodeCodes(6, 2)
Dim BudstrSelCri As String

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()

End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    cmdPrint.Enabled = False
    cmdView.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    cmdPrint.Enabled = True
    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim x%
Dim Y%
ReDim glbCode_Snap(6)

glbOnTop = Me.name
Screen.MousePointer = HOURGLASS

'Ticket #26726 Franks 06/17/2015 - open it for all
'If glbWFC Then 'Ticket #25911 Franks 01/27/2015
    clpJOB.TextBoxWidth = 1265
'End If

Me.Show

Screen.MousePointer = HOURGLASS

Call setRptCaption(Me)

lblGL.Caption = lStr(lblGL.Caption)

If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

Call INI_Controls(Me)

If glbMulti Then lblUnion.ForeColor = &HC000C0: lblPT.ForeColor = &HC000C0

Screen.MousePointer = DEFAULT

End Sub

Private Function CriCheck()
Dim x%, intDSet%, I
Dim BudgetFlag As Boolean, NonBudFlag As Boolean
Dim xMsg As String

CriCheck = False

BudgetFlag = False: NonBudFlag = False

If Not clpDiv.ListChecker Then
'If Len(clpDiv) > 0 And clpDiv.Caption = "Unassigned" Then
    'MsgBox lStr("If Division Entered - it must be known")
    'clpDiv.SetFocus
    Exit Function
End If

If Len(clpDiv) > 0 Then
    BudgetFlag = True
End If

If Not clpDept.ListChecker Then
'If Len(clpDept) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

If Len(clpDept) > 0 Then
    BudgetFlag = True
End If

If Len(clpGLNum) > 0 And clpGLNum.Caption = "Unassigned" Then
    MsgBox "If G/L Code Entered - it must be known"
    clpGLNum.SetFocus
    Exit Function
End If

If Len(clpGLNum) > 0 Then
    BudgetFlag = True
End If

If Not clpPT.ListChecker Then
'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Len(clpPT) > 0 Then
    NonBudFlag = True
End If

If Len(clpJOB) > 0 And clpJOB.Caption = "Unassigned" Then
    MsgBox "If code entered it must be known"
    clpJOB.SetFocus
    Exit Function
End If

If Len(clpJOB) > 0 Then
    NonBudFlag = True
End If

If Len(elpEEID) > 0 Then
    NonBudFlag = True
End If

For x% = 0 To 6
    If Not clpCode(x).ListChecker Then Exit Function

    If Len(clpCode(x%)) > 0 Then
        NonBudFlag = True
    End If
Next x%

If NonBudFlag And Not BudgetFlag Then
        xMsg = "If Selection Criteria entered is using data in the Purple or Black," & Chr(10)
        xMsg = xMsg & "at least one Red Criteria must be entered"
        'MsgBox "If Selection Criteria entered is using data in the Purple or Black, at least one Red Criteria must be entered"
        MsgBox xMsg
        clpDiv.SetFocus
        Exit Function
End If
BudstrSelCri = ""
intDSet% = False

CriCheck = True

End Function

Private Function Cri_SetAll()
Dim x%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept)
If glbMulti Then
glbstrSelCri = Replace(glbstrSelCri, "HREMP.ED_DEPTNO", "HR_JOB_HISTORY.JH_DEPTNO")
End If
Call Cri_Div
Call Cri_DEPTNO
Call Cri_GLNO
If Len(BudstrSelCri) > 0 Then
    BudstrSelCri = "(1 = 1) " & BudstrSelCri
End If
Call Cri_Position

For x% = 0 To 6
    Call Cri_Code(x%)
Next
Call Cri_PT
Call Cri_EE
Call Cri_FTDates

glbstrSelCri = Replace(Replace(glbstrSelCri, "{", "("), "}", ")")
Call PLANRPT_WRK

Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZPOE.rpt"
Me.vbxCrystal.WindowTitle = "Plan of Establishment Report"

glbstrSelCri = "{HRPOERPT.JH_WRKEMP}='" & glbUserID & "'"
Me.vbxCrystal.SelectionFormula = glbstrSelCri
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 11
        If x% <> 6 Then
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Else
        Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
        End If
    Next x%
End If

''If glbNoNONE Then
''    'Modified by Franks on Feb 12,2002 for missing data when the user's union set to "-NON"
''    'Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND ({HREMP.ED_ORG }<> 'NONE')"
''    Me.vbxCrystal.SelectionFormula = "(" & glbstrSelCri & " ) AND (isnull({HREMP.ED_ORG }) OR {HREMP.ED_ORG }<> 'NONE')"
''    'Modified by Franks on Feb 12,2002
''Else
''    Me.vbxCrystal.SelectionFormula = glbstrSelCri
''End If

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

Private Sub PLANRPT_WRK()
Dim rsWRK As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim rsBUD As New ADODB.Recordset
Dim rsWR1 As New ADODB.Recordset
Dim rsWR2 As New ADODB.Recordset
Dim rsJCODE As New ADODB.Recordset
Dim rsPosCtrl As New ADODB.Recordset
Dim SQLQ, xNUMBER, xEMP, xTEMPDATE, xINT, I
Dim xJob, xDeptno, xGLNO, xID, xStr, xDiv, xINTorg, xPosCtrl
Dim xNum, RecNum
Dim Langs 'George Apr 4,2006 #10574
    'DELETE ALL RECORDS FROM TEMP TABLE
    SQLQ = "DELETE FROM HRPOERPT WHERE JH_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.CursorLocation = adUseClient
    gdbAdoIhr001.BeginTrans
    If glbSQL Or glbOracle Then
        gdbAdoIhr001.Execute SQLQ
    Else
         gdbAdoIhr001W.Execute SQLQ
    End If
    gdbAdoIhr001.CommitTrans
    
    SQLQ = "DELETE FROM HRPOETMP WHERE JH_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001.BeginTrans
    If glbSQL Or glbOracle Then
        gdbAdoIhr001.Execute SQLQ
    Else
         gdbAdoIhr001W.Execute SQLQ
    End If
    gdbAdoIhr001.CommitTrans
    glbstrSelCri = Replace(Replace(glbstrSelCri, "[", "("), "]", ")")
    If glbOracle Then
        'SQLQ = "SELECT HR_JOB_HISTORY.*, HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_LANG1,HREMP.ED_LANG2, HREMP.ED_DOH " 'George Apr 4,2006 #10574
        SQLQ = "SELECT HR_JOB_HISTORY.*, HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DOH "
        If Not glbMulti Then
            SQLQ = SQLQ & ",HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_GLNO,ED_EMP,ED_SECTION "
        End If
        SQLQ = SQLQ & "FROM HR_JOB_HISTORY,HREMP,HRJOB "
        SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 "
        SQLQ = SQLQ & "AND " & glbstrSelCri & " "
        SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & "AND HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    Else
        'SQLQ = "SELECT HR_JOB_HISTORY.*, HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_LANG1,HREMP.ED_LANG2, HREMP.ED_DOH "'George Apr 4,2006 #10574
        SQLQ = "SELECT HR_JOB_HISTORY.*, HREMP.ED_SURNAME,HREMP.ED_FNAME,HREMP.ED_DOH "
        If Not glbMulti Then
            SQLQ = SQLQ & ",HREMP.ED_DIV,HREMP.ED_DEPTNO,HREMP.ED_GLNO,ED_EMP,ED_SECTION "
        End If
        SQLQ = SQLQ & "FROM (HR_JOB_HISTORY LEFT JOIN HREMP ON HR_JOB_HISTORY.JH_EMPNBR = HREMP.ED_EMPNBR) "
        SQLQ = SQLQ & "LEFT JOIN HRJOB ON HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
        SQLQ = SQLQ & "WHERE HR_JOB_HISTORY.JH_CURRENT <> 0 "
        SQLQ = SQLQ & "AND " & glbstrSelCri & " "
    End If
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If glbSQL Or glbOracle Then
        rsWRK.Open "SELECT * FROM HRPOERPT WHERE JH_WRKEMP = '" & glbUserID & "' ", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        rsWRK.Open "SELECT * FROM HRPOERPT WHERE JH_WRKEMP = '" & glbUserID & "' ", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    End If
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = "Please wait..."
    If Not rsJOB.EOF Then
        rsJOB.MoveLast
        rsJOB.MoveFirst
        RecNum = rsJOB.RecordCount
        xNum = 0
    End If
    Do While Not rsJOB.EOF
        MDIMain.panHelp(0).FloodPercent = (xNum / RecNum) * 100
        xNum = xNum + 1
        xNUMBER = rsJOB("JH_EMPNBR")
        rsWRK.AddNew
        rsWRK("JH_COMPNO") = "001"
        If glbMulti Then
            rsWRK("JH_DIV") = rsJOB("JH_DIV")
            rsWRK("JH_DEPTNO") = rsJOB("JH_DEPTNO")
            rsWRK("JH_GLNO") = rsJOB("JH_GLNO") '
            rsWRK("JH_EMP") = rsJOB("JH_EMP")
            rsWRK("JH_SECTION") = rsJOB("JH_SECTION")
        Else
            rsWRK("JH_DIV") = rsJOB("ED_DIV")
            rsWRK("JH_DEPTNO") = rsJOB("ED_DEPTNO")
            rsWRK("JH_GLNO") = rsJOB("ED_GLNO") '
            rsWRK("JH_EMP") = rsJOB("ED_EMP")
            rsWRK("JH_SECTION") = rsJOB("ED_SECTION")
        End If
        rsWRK("JH_EMP_TABL") = "EDEM" 'rsJOB("JH_EMP_TABL")
        rsWRK("JH_SECTION_TABL") = "EDSE"
        rsWRK("JH_JOB") = rsJOB("JH_JOB")
        rsWRK("JH_EMPNBR") = xNUMBER
        rsWRK("JH_NAME") = rsJOB("ED_SURNAME") & ", " & rsJOB("ED_FNAME")
        rsWRK("JH_SDATE") = rsJOB("JH_SDATE")
        rsWRK("JH_CURRENT") = -1
        rsWRK("JH_JREASON") = rsJOB("JH_JREASON")
        rsWRK("JH_REPTAU") = rsJOB("JH_REPTAU")
        rsWRK("JH_DHRS") = rsJOB("JH_DHRS")
        rsWRK("JH_WHRS") = rsJOB("JH_WHRS")
        rsWRK("JH_PHRS") = rsJOB("JH_PHRS")
        rsWRK("JH_SHIFT") = rsJOB("JH_SHIFT")
        'rsWRK("JH_FTENUM") = rsJOB("JH_FTENUM")
        rsWRK("JH_FTEHRS") = rsJOB("JH_FTEHRS")
        rsWRK("JH_ORG") = rsJOB("JH_ORG")
        rsWRK("JH_ORG_TABL") = "EDOR" 'rsJOB("JH_ORG_TABL")
        rsWRK("JH_PT") = rsJOB("JH_PT")
        rsWRK("JH_COMMENT") = rsJOB("JH_COMMENT")
        rsWRK("JH_REPTAU2") = rsJOB("JH_REPTAU2")
        rsWRK("JH_REPTAU3") = rsJOB("JH_REPTAU3")
        rsWRK("JH_REPTAU3") = rsJOB("JH_REPTAU3")
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xNUMBER & " "
        SQLQ = SQLQ & "AND SH_CURRENT <> 0"
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsSal.EOF Then '
            rsWRK("JH_SALARY") = rsSal("SH_SALARY")
            If IsDate(rsSal("SH_NEXTDAT")) Then
                rsWRK("JH_NEXTDAT") = rsSal("SH_NEXTDAT")
            End If
        End If
        rsSal.Close
        ''Add Position Control # here Begin Open recordset for HRJOBBUD
        'SQLQ = "SELECT * FROM HRJOBBUD WHERE "
        'SQLQ = SQLQ & "JG_CODE = '" & rsJOB("JH_JOB") & "' "
        'If glbMulti Then
        '    SQLQ = SQLQ & "AND JG_DEPTNO = '" & rsJOB("JH_DEPTNO") & "' "
        '    If Not IsNull(rsJOB("JH_DIV")) Then
        '        SQLQ = SQLQ & "AND JG_DIV = '" & rsJOB("JH_DIV") & "' "
        '    End If
        '    If Not IsNull(rsJOB("JH_GLNO")) Then
        '        SQLQ = SQLQ & "AND JG_GLNO = '" & rsJOB("JH_GLNO") & "' "
        '    End If
        'Else
        '    SQLQ = SQLQ & "AND JG_DEPTNO = '" & rsJOB("ED_DEPTNO") & "' "
        '    If Not IsNull(rsJOB("ED_DIV")) Then
        '        SQLQ = SQLQ & "AND JG_DIV = '" & rsJOB("ED_DIV") & "' "
        '    End If
        '    If Not IsNull(rsJOB("ED_GLNO")) Then
        '        SQLQ = SQLQ & "AND JG_GLNO = '" & rsJOB("ED_GLNO") & "' "
        '    End If
        'End If
        'If rsPosCtrl.State <> 0 Then rsPosCtrl.Close
        'rsPosCtrl.Open SQLQ, gdbAdoIhr001, adOpenStatic
        'If Not rsPosCtrl.EOF Then
        '    rsWRK("JH_POSCTRLNO") = rsPosCtrl("JG_POSCTRLNO")
        'End If
        'rsPosCtrl.Close
        ''Add Position Control # here - End
        rsWRK("JH_LANG_TABL") = "EDL1"
        'rsWrk("JH_LANG") = rsJOB("ED_LANG1") 'George Apr 4,2006 #10574
        Langs = Split(getLanguage(xNUMBER), "|")
        If Langs(0) <> "NoLang1" Then rsWRK("JH_LANG") = Langs(0) '0 is for ED_Lang1, 1 is ED_Lang2
        'Comment by Frank on Oct 28,2002, don't hard code "BOTH", get code from table
        'If Not IsNull(rsJOB("ED_LANG1")) And Not IsNull(rsJOB("ED_LANG2")) Then
        '    If rsJOB("ED_LANG1") <> rsJOB("ED_LANG2") Then
        '        If InStr(1, "ENGL,FREN", rsJOB("ED_LANG1")) > 0 Then
        '            If InStr(1, "ENGL,FREN", rsJOB("ED_LANG2")) > 0 Then
        '                rsWRK("JH_LANG") = "BOTH"
        '            End If
        '        End If
        '    End If
        'End If
        If glbMulti Then
            If Not IsNull(rsJOB("JH_EMP")) Then
                If InStr(1, ",CAS,CONT,TA,", "," & rsJOB("JH_EMP") & ",") > 0 Then
                    rsWRK("JH_CCTEMP") = rsJOB("JH_EMP")
                    If rsJOB("JH_EMP") = "TA" Then
                        rsWRK("JH_TEMPANN") = rsJOB("JH_SDATE")
                    End If
                    rsWRK("JH_CCTNBR") = xNUMBER
                    rsWRK("JH_DELFLAG") = "2"
                    'rsWRK("JH_NAME") = "VACANT"
                Else
                    rsWRK("JH_DELFLAG") = "1"
                    rsWRK("JH_FTENUM") = rsJOB("JH_FTENUM")
                End If
            Else
                rsWRK("JH_DELFLAG") = "1"
            End If
        Else
            If Not IsNull(rsJOB("ED_EMP")) Then
                If InStr(1, ",CAS,CONT,TA,", "," & rsJOB("ED_EMP") & ",") > 0 Then
                    rsWRK("JH_CCTEMP") = rsJOB("ED_EMP")
                    If rsJOB("ED_EMP") = "TA" Then
                        rsWRK("JH_TEMPANN") = rsJOB("JH_SDATE")
                    End If
                    rsWRK("JH_CCTNBR") = xNUMBER
                    rsWRK("JH_DELFLAG") = "2"
                    'rsWRK("JH_NAME") = "VACANT"
                Else
                    rsWRK("JH_DELFLAG") = "1"
                    rsWRK("JH_FTENUM") = rsJOB("JH_FTENUM")
                End If
            Else
                rsWRK("JH_DELFLAG") = "1"
            End If
        End If
        rsWRK("JH_WRKEMP") = glbUserID
        rsWRK.Update
        rsJOB.MoveNext
    Loop
    rsWRK.Close
    rsJOB.Close

    'Data process
    ''Delete all records which JH_DELFLAG = 1
    'SQLQ = "DELETE FROM HRPOERPT WHERE JH_DELFLAG = '1'"
    'gdbAdoIhr001.BeginTrans
    'gdbAdoIhr001.Execute SQLQ
    'gdbAdoIhr001.CommitTrans
    
    
    'Check Budgeted Position Table,
    'If there are other vacancy FTE for this position, print these records in the report
    If glbSQL Or glbOracle Then
        rsWR1.Open "SELECT * FROM HRPOETMP WHERE JH_WRKEMP = '" & glbUserID & "' ", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        rsWR1.Open "SELECT * FROM HRPOETMP WHERE JH_WRKEMP = '" & glbUserID & "' ", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    End If
    
    SQLQ = "SELECT * FROM HRJOBBUD "
    If Len(BudstrSelCri) > 0 Then
        SQLQ = SQLQ & "WHERE " & BudstrSelCri
    End If
    If rsWRK.State <> 0 Then rsWRK.Close
    rsWRK.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsWRK.EOF Then
        rsWRK.MoveLast
        rsWRK.MoveFirst
        RecNum = rsWRK.RecordCount
        xNum = 0
    End If
    Do While Not rsWRK.EOF
        MDIMain.panHelp(0).FloodPercent = (xNum / RecNum) * 100
        xNum = xNum + 1
        xJob = rsWRK("JG_CODE")
        If Not IsNull(rsWRK("JG_DIV")) Then
            xDiv = rsWRK("JG_DIV")
        Else
            xDiv = ""
        End If
        xDeptno = rsWRK("JG_DEPTNO")
        xGLNO = rsWRK("JG_GLNO")
        'If IsNull(rsWRK("JG_POSCTRLNO")) Then
        '    xPosCtrl = ""
        'Else
        '    xPosCtrl = rsWRK("JG_POSCTRLNO")
        'End If
        xINT = rsWRK("JG_FTENUMVACN"): xINTorg = xINT
        If xINT > 0 Then
            'If xINT > Int(xINT) Then
            '    xINT = Int(xINT) + 1
            'End If
            'For I = 1 To xINT
                rsWR1.AddNew
                rsWR1("JH_COMPNO") = "001"
                'rsWR1("JH_POSCTRLNO") = xPosCtrl
                rsWR1("JH_DIV") = xDiv 'rsWRK("JH_DIV")
                rsWR1("JH_DEPTNO") = xDeptno 'rsWRK("JH_DEPTNO")
                rsWR1("JH_GLNO") = xGLNO 'rsWRK("JH_GLNO")
                rsWR1("JH_JOB") = xJob 'rsWRK("JH_JOB")
                rsWR1("JH_FTENUM") = xINTorg
                SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xJob & "' "
                rsJCODE.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsJCODE.EOF Then
                    If Not IsNull(rsJCODE("JB_ORG")) Then
                        rsWR1("JH_ORG") = rsJCODE("JB_ORG")
                    End If
                End If
                rsJCODE.Close
                rsWR1("JH_WRKEMP") = glbUserID
                rsWR1.Update
            'Next I
        End If
        rsWRK.MoveNext
    Loop
    rsWRK.Close
    If glbSQL Or glbOracle Then
        rsWRK.Open "SELECT * FROM HRPOERPT WHERE JH_WRKEMP = '" & glbUserID & "' ", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        rsWRK.Open "SELECT * FROM HRPOERPT WHERE JH_WRKEMP = '" & glbUserID & "' ", gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    End If
    If Not rsWR1.EOF And Not rsWR1.BOF Then
        rsWR1.MoveLast
        rsWR1.MoveFirst
        RecNum = rsWR1.RecordCount
        xNum = 0
        Do While Not rsWR1.EOF
            'MDIMain.panHelp(0).FloodPercent = (xNum / RecNum) * 100
            xNum = xNum + 1
            rsWRK.AddNew
            rsWRK("JH_COMPNO") = "001"
            'rsWRK("JH_POSCTRLNO") = rsWR1("JH_POSCTRLNO")
            rsWRK("JH_DIV") = rsWR1("JH_DIV")
            rsWRK("JH_DEPTNO") = rsWR1("JH_DEPTNO")
            rsWRK("JH_GLNO") = rsWR1("JH_GLNO")
            rsWRK("JH_JOB") = rsWR1("JH_JOB")
            rsWRK("JH_ORG") = rsWR1("JH_ORG")
            rsWRK("JH_FTENUM") = rsWR1("JH_FTENUM")
            rsWRK("JH_NAME") = "VACANT"
            rsWRK("JH_DELFLAG") = "3"
            rsWRK("JH_WRKEMP") = glbUserID
            rsWRK.Update
            rsWR1.MoveNext
        Loop
    End If
    rsWR1.Close
    rsWRK.Close
    If glbOracle Then
        gdbAdoIhr001.CursorLocation = adUseServer
    End If
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = ""
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%)) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If glbMulti Then
        If intIdx% = 1 Then strCd$ = "HR_JOB_HISTORY.JH_ORG"
        If intIdx% = 2 Then strCd$ = "HR_JOB_HISTORY.JH_EMP"
    Else
        If intIdx% = 1 Then strCd$ = "HREMP.ED_ORG"
        If intIdx% = 2 Then strCd$ = "HREMP.ED_EMP"
    End If
    If intIdx% = 3 Then strCd$ = "HRJOB.JB_GRPCD"
    If intIdx% = 4 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 6 Then strCd$ = "HREMP.ED_SECTION"  'Lucy June 30, 2000
        CodeCri = "({" & strCd$ & "} in  ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "'))"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv & clpCode(intIdx%) & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%) & "') )"
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
Dim DivCri As String, DivCriB
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Len(clpDiv.Text) > 0 Then
    If glbMulti Then
        DivCri = "({HR_JOB_HISTORY.JH_DIV} in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
    Else
        DivCri = "({HREMP.ED_DIV} in ('" & Replace(clpDiv.Text, ",", "','") & "'))"
    End If
    DivCriB = "(JG_DIV = '" & clpDiv.Text & "')"
End If

If Len(DivCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivCri
        BudstrSelCri = DivCriB
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivCri
        BudstrSelCri = BudstrSelCri & " AND " & DivCriB
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_DEPTNO()
Dim DEPTGLNO As String, DEPTGLNOB

If Len(clpDept.Text) > 0 Then
    DEPTGLNO = "(JG_DEPTNO in ('" & Replace(clpDept.Text, ",", "','") & "'))"
End If

If Len(DEPTGLNO) >= 1 Then
    If Not glbiOneWhere Then
        BudstrSelCri = DEPTGLNO
    Else
        BudstrSelCri = BudstrSelCri & " AND " & DEPTGLNO
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_GLNO()
Dim DivGLNO As String, DivGLNOB
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Len(clpGLNum.Text) > 0 Then
    If glbMulti Then
        DivGLNO = "({HR_JOB_HISTORY.JH_GLNO} = '" & clpGLNum.Text & "')"
    Else
        DivGLNO = "({HREMP.ED_GLNO} = '" & clpGLNum.Text & "')"
    End If
    DivGLNOB = "(JG_GLNO = '" & clpGLNum.Text & "')"
End If

If Len(DivGLNO) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = DivGLNO
        BudstrSelCri = DivGLNOB
    Else
        glbstrSelCri = glbstrSelCri & " AND " & DivGLNO
        BudstrSelCri = BudstrSelCri & " AND " & DivGLNOB
    End If
    glbiOneWhere = True
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    EECri = "{HREMP.ED_EMPNBR} IN (" & getEmpnbr(elpEEID) & ") "
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
'Dim TempCri As String
'Dim dtYYY%, dtMM%, dtDD%
'Dim X%, strCriNam$
'
'strCriNam$ = "HR_JOB_HISTORY.JH_SDATE"
'
'DoCrit:
'If Len(txtDateRange(0).Text) > 0 And Len(txtDateRange(1).Text) > 0 Then
'    TempCri = "({" & strCriNam$ & "} "
'    dtYYY% = Year(txtDateRange(0).Text)
'    dtMM% = Month(txtDateRange(0).Text)
'    dtDD% = Day(txtDateRange(0).Text)
'    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
'    dtYYY% = Year(txtDateRange(1).Text)
'    dtMM% = Month(txtDateRange(1).Text)
'    dtDD% = Day(txtDateRange(1).Text)
'    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'    GoTo Cri_FTDatst
'End If
'
'For X% = 0 To 1
'    If Len(txtDateRange(0).Text) > 0 Then
'        TempCri = "({" & strCriNam$ & "} "
'        If X% = 0 Then
'            TempCri = TempCri & " >= "
'        Else
'            TempCri = TempCri & " <= "
'        End If
'        dtYYY% = Year(txtDateRange(0).Text)
'        dtMM% = Month(txtDateRange(0).Text)
'        dtDD% = Day(txtDateRange(0).Text)
'        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'        GoTo Cri_FTDatst
'    End If
'Next X%
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

Private Sub Cri_PT()
Dim EECri As String, OneSet%, x%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HR_JOB_HISTORY.JH_PT} in ('" & Replace(clpPT.Text, ",", "','") & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If

glbiOneWhere = True

End Sub

Private Sub Cri_Position()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String
If Len(clpJOB.Text) <= 0 Then Exit Sub
PosCri = "({HR_JOB_HISTORY.JH_JOB} = '" & clpJOB.Text & "')"
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
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
