VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmWorkForce 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employer Workforce Survey"
   ClientHeight    =   7005
   ClientLeft      =   315
   ClientTop       =   1020
   ClientWidth     =   11730
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7005
   ScaleWidth      =   11730
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkShowEmp 
      Caption         =   "Show Employee Details"
      Height          =   255
      Left            =   320
      TabIndex        =   9
      Top             =   3600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtSurveyCompleted 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2300
      MaxLength       =   1
      TabIndex        =   7
      Tag             =   "01-Enter Y / N"
      Top             =   2760
      Width           =   375
   End
   Begin INFOHR_Controls.CodeLookup clpPlanNbr 
      Height          =   285
      Left            =   1980
      TabIndex        =   6
      Tag             =   "01-Enter Plan Number"
      Top             =   2412
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   10
      LookupType      =   7
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1980
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   677
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
      Left            =   1980
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   330
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
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Top             =   1024
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
      Left            =   1980
      TabIndex        =   3
      Top             =   1371
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
      Left            =   1980
      TabIndex        =   4
      Top             =   1718
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1980
      TabIndex        =   5
      Tag             =   "10-Enter Employee Number"
      Top             =   2065
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6840
      Top             =   4080
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
   Begin INFOHR_Controls.CodeLookup clpNationalClass 
      Height          =   285
      Left            =   1980
      TabIndex        =   8
      Tag             =   "00-National Occupation Classification -Code"
      Top             =   3120
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   6
   End
   Begin VB.Label lblNOC 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N.O.C. Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   19
      Top             =   3165
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblSurveyCompl 
      Appearance      =   0  'Flat
      Caption         =   "Survey Completed"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   18
      Top             =   2805
      Width           =   1695
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   17
      Top             =   1763
      Width           =   630
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
      Left            =   300
      TabIndex        =   16
      Top             =   2110
      Width           =   1290
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
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblPlanNo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan Number"
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
      Left            =   300
      TabIndex        =   14
      Top             =   2457
      Width           =   1095
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   1416
      Width           =   450
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union Code"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   1069
      Width           =   840
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   722
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
      Left            =   300
      TabIndex        =   10
      Top             =   375
      Width           =   555
   End
End
Attribute VB_Name = "frmWorkForce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PlanNo_Snap As New ADODB.Recordset
Dim LastlastID, LastlastNme, LastFirstNme, xTxtEEID, xLblEEName

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Employer Workforce Survey Criteria", Me) Then Exit Sub
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

Exit Sub

PrntErr:
MsgBox "Error Printing - check your Windows Printer setup"

Call RollBack  '08June99 js

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Sub cmdView_Click()
Dim X%
Dim strWHand As String

On Error GoTo CRW_Err

If CriCheck() Then
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
End If

Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
Call RollBack  '08June99 js

End Sub

'Private Sub cmdView_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub CR_PLAN_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo Plan_Err

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HRPARCOP "

If PlanNo_Snap.State <> 0 Then Exit Sub
PlanNo_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

If PlanNo_Snap.EOF And PlanNo_Snap.BOF Then
    Msg = "No Plan Number descriptions found" & Chr(10)
    Msg = Msg & "You will require authority to add one to continue"
    MsgBox Msg
    Exit Sub
Else
  PlanNo_Snap.MoveFirst
End If

Screen.MousePointer = DEFAULT

Exit Sub

Plan_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Plan", "HRPARCOP", "SELECT")
Call RollBack '08June99 js

End Sub

Private Sub Cri_Assoc()
Dim EECri As String

If Len(clpCode(1).Text) > 0 Then
    'Ticket #24906 - Vitalaire report changes
    'EECri = "{HREMPEQU.EQ_ORG} = '" & clpCode(1).Text & "' "
    EECri = "{HREMPEQU.EQ_ORG} in ['" & Replace(clpCode(1).Text, ",", "','") & "']"
End If

If Len(EECri) >= 1 Then
    Call fnOneWhere(EECri)   '08June99 js
End If

End Sub

Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

If Len(clpDiv.Text) > 0 Then
    'Ticket #24906 - Vitalaire report changes
    'DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
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

Private Sub Cri_PlanNo()
Dim EECri As String, OneSet%, X%

If Len(clpPlanNbr.Text) < 1 Then Exit Sub

EECri = "{HRPARCOP.PP_PLAN}= '" & clpPlanNbr.Text & "'"

Call fnOneWhere(EECri)  '08June99 js

End Sub

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub

'Ticket #24906 - Vitalaire report changes
'EECri = "{HREMPEQU.EQ_EEPT}= '" & clpPT.Text & "'"
EECri = "{HREMPEQU.EQ_EEPT} in ['" & Replace(clpPT.Text, ",", "','") & "']"

Call fnOneWhere(EECri)  '08June99 js

End Sub

Private Sub Cri_SurveyComplete()
Dim EECri As String, OneSet%, X%

If Len(txtSurveyCompleted.Text) < 1 Then Exit Sub

EECri = "{HREMPEQU.EQ_SURCOMP}= '" & txtSurveyCompleted.Text & "'"

Call fnOneWhere(EECri)  '08June99 js

End Sub

Private Sub Cri_NOC_EEOG()
Dim EECri As String, OneSet%, X%

If Len(clpNationalClass.Text) < 1 Then Exit Sub

EECri = "{HREMPEQU.EQ_NOGC}= '" & clpNationalClass.Text & "'"

Call fnOneWhere(EECri)  '08June99 js

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
Call glbCri_DeptUN(clpDept.Text)
Call Cri_PlanNo
Call Cri_Org
Call Cri_Div
Call Cri_Assoc
Call Cri_Status
Call Cri_PT
Call Cri_SurveyComplete
Call Cri_NOC_EEOG

' report name
If frmWorkForce.Caption = "Completed Workforce Surveys" Then
    strRName$ = glbIHRREPORTS & "rzsurv.rpt"
    Me.vbxCrystal.WindowTitle = "Completed Workforce Surveys"
ElseIf frmWorkForce.Caption = "Employment Status Analysis" And chkShowEmp.Value <> 1 Then   'Ticket #24906 - Vitalaire report changes
    strRName$ = glbIHRREPORTS & "rzeestat.rpt"
    Me.vbxCrystal.WindowTitle = "Employment Status Analysis"
'Ticket #24906 - Vitalaire report changes
ElseIf frmWorkForce.Caption = "Employment Status Analysis" And chkShowEmp.Value = 1 Then
    strRName$ = glbIHRREPORTS & "rzeestatE.rpt"
    Me.vbxCrystal.WindowTitle = "Employment Status Analysis"
ElseIf frmWorkForce.Caption = "Occupational Group Analysis" Then
    strRName$ = glbIHRREPORTS & "rzeeoccu.rpt"
    Me.vbxCrystal.WindowTitle = "Occupational Group Analysis"
Else
    strRName$ = glbIHRREPORTS & "rzeework.rpt"
    Me.vbxCrystal.WindowTitle = "Employer Workforce Surveys"
End If

Me.vbxCrystal.ReportFileName = strRName$

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL

Cri_SetAll = True
Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Cri_SetAll = False

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Time", "Comp Report", "Select")
Call RollBack

End Function

Private Sub Cri_Status()
Dim EECri As String

If Len(clpCode(2).Text) > 0 Then
    'Ticket #24906 - Vitalaire report changes
    'EECri = "{HREMP.ED_EMP} = '" & clpCode(2).Text & "' "
    EECri = "{HREMP.ED_EMP} in ['" & Replace(clpCode(2).Text, ",", "','") & "']"
End If

If Len(EECri) >= 1 Then
    Call fnOneWhere(EECri)  '08June99 js
End If

End Sub

Private Sub Cri_Org()
Dim EECri As String

If Len(clpCode(1).Text) > 0 Then
    'Ticket #24906 - Vitalaire report changes
    'EECri = "{HREMP.ED_ORG} = '" & clpCode(1).Text & "' "
    EECri = "{HREMP.ED_ORG} in ['" & Replace(clpCode(1).Text, ",", "','") & "']"
End If

If Len(EECri) >= 1 Then
    Call fnOneWhere(EECri)  '08June99 js
End If

End Sub

Private Function CriCheck()
Dim X%

CriCheck = False

'Ticket #24906 - Vitalaire report changes
If Not clpDiv.ListChecker Then
'If Len(clpDiv) > 0 And clpDiv.Caption = "Unassigned" Then
'    MsgBox lStr("If Division Entered - it must be known")
'    clpDiv.SetFocus
    Exit Function
End If

'Ticket #24906 - Vitalaire report changes
If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
'    MsgBox lStr("If Department Entered - it must be known")
'    clpDept.SetFocus
    Exit Function
End If

If Not clpCode(1).ListChecker Then Exit Function

'Ticket #24906 - Vitalaire report changes
If Not clpCode(2).ListChecker Then
'If Len(clpCode(2).Text) > 0 And clpCode(2).Caption = "Unassigned" Then
'    MsgBox "If Status Code is entered it must be known"
'    clpCode(2).SetFocus
    Exit Function
End If

'Ticket #24906 - Vitalaire report changes
If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
'    MsgBox lStr("Category code must be valid")
'    clpPT.SetFocus
    Exit Function
End If

If clpPlanNbr.Text = "" Then
    MsgBox "Plan Number is a required entry!"
    clpPlanNbr.SetFocus
    Exit Function
End If

If clpPlanNbr.Caption = "Unassigned" Then
    MsgBox "Please enter a valid Plan Number!"
    clpPlanNbr.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

CriCheck = True

End Function

Private Sub Form_Activate()
glbOnTop = "FRMWORKFORCE"
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

glbOnTop = "FRMWORKFORCE"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Screen.MousePointer = HOURGLASS

Call setRptCaption(Me)

'Ticket #19537
If glbCompSerial = "S/N - 2279W" Then
    lblNOC.Visible = True
    clpNationalClass.Visible = True
    lblNOC.Caption = "EEOG - N.O.C. Code"
    clpNationalClass.Tag = "00-EEOG -N.O.C. Code"
End If

If Not gSec_Upd_EmploymentEQT Then     'May99 js
 '   cmdView.Enabled = False             '
    clpCode(1).Enabled = False          '
    clpCode(2).Enabled = False          '
     clpDept.Enabled = False             '
     clpDiv.Enabled = False              '
     clpPlanNbr.Enabled = False          '
     clpPT.Enabled = False               '
End If                                  '

'Ticket #24906 - Vitalaire report changes
If frmWorkForce.Caption = "Employment Status Analysis" Then
    chkShowEmp.Visible = True
Else
    chkShowEmp.Visible = False
End If

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
End Sub

Private Sub Plan_Desc()
Dim SQLQ As String

On Error GoTo PlanNo_Err

 clpPlanNbr.ShowDescription = True
 clpPlanNbr.Caption = "Unassigned"

If Len(clpPlanNbr.Text) > 0 Then
    SQLQ = "PP_PLAN = '" & clpPlanNbr.Text & "'"
    PlanNo_Snap.Requery
    PlanNo_Snap.Find SQLQ
    If Not PlanNo_Snap.EOF Then
        clpPlanNbr.Caption = PlanNo_Snap("PP_DESC")
        clpPlanNbr.ShowDescription = True
    End If
End If

Exit Sub

PlanNo_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Plan Snap", "Plan Number", "SELECT")
Call RollBack '08June99 js

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

Private Function fnOneWhere(EECri As String)    '08June99 js
If Len(glbstrSelCri) > 0 Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
End Function

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

Private Sub txtSurveyCompleted_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSurveyCompleted_LostFocus()
    'Hemu - 07/02/2003 Begin - Jerry suggested
    txtSurveyCompleted.Text = UCase(txtSurveyCompleted.Text)
    'Hemu - 07/02/2003 End
End Sub

