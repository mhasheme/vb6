VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRComPlan 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Compensation Plan Report"
   ClientHeight    =   5250
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   8295
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
   ScaleHeight     =   5250
   ScaleWidth      =   8295
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkWithAct 
      Caption         =   "With Actuals"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   4
      Tag             =   "With Actuals Layout"
      Top             =   2070
      Width           =   2475
   End
   Begin VB.CheckBox chkCommPlan 
      Caption         =   "Show Commission Plan"
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
      Left            =   360
      TabIndex        =   3
      Tag             =   "Show Commission Plan"
      Top             =   1800
      Width           =   2535
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   4590
      Width           =   8295
      _Version        =   65536
      _ExtentX        =   14631
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   150
         TabIndex        =   15
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton cmdView 
         Appearance      =   0  'Flat
         Caption         =   "&View"
         Height          =   375
         Left            =   1110
         TabIndex        =   16
         Tag             =   "View this report (with option to print)"
         Top             =   150
         Width           =   855
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   2100
         TabIndex        =   17
         Tag             =   "Print the report"
         Top             =   150
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   3510
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
         WindowShowPrintSetupBtn=   -1  'True
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
      Index           =   1
      Left            =   2005
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Second level of grouping records"
      Top             =   4200
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
      Left            =   2005
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "First Level of grouping records"
      Top             =   3885
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   690
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Tag             =   "10-Enter Employee Number"
      Top             =   1020
      Width           =   7995
      _ExtentX        =   14102
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7675
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
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
      Left            =   120
      TabIndex        =   13
      Top             =   3675
      Width           =   1575
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
      Left            =   330
      TabIndex        =   12
      Top             =   4230
      Width           =   660
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
      Left            =   330
      TabIndex        =   11
      Top             =   3915
      Width           =   885
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
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   1575
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
      Left            =   330
      TabIndex        =   9
      Top             =   1050
      Width           =   1290
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
      Left            =   330
      TabIndex        =   8
      Top             =   750
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
      Left            =   330
      TabIndex        =   5
      Top             =   420
      Width           =   555
   End
End
Attribute VB_Name = "frmRComPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc
Dim xTxtEEID, xLblEEName

Private Sub chkCommPlan_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub chkWithAct_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr
 
If CriCheck() Then

    If Not PrtForm("Email Address Report Criteria", Me) Then Exit Sub

    cmdPrint.Enabled = False
    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    cmdPrint.Enabled = True
    cmdView.Enabled = True
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub cmdPrint_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    cmdPrint.Enabled = False
    cmdView.Enabled = False
    Screen.MousePointer = HOURGLASS
    
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
    cmdPrint.Enabled = True
    cmdView.Enabled = True
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

Private Sub cmdView_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem "Employee Name"
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    comGroup(1).AddItem "Employee Name"
    comGroup(1).ListIndex = 0

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

If Len(getEmpnbr(elpEEID.Text)) > 0 Then
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



Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_EE
If glbCompSerial = "S/N - 2291W" Then
    If chkWithAct = 0 Then
        If chkCommPlan = 0 Then
            strRName$ = glbIHRREPORTS & "sn2291-1.rpt"
        Else
            strRName$ = glbIHRREPORTS & "sn2291-2.rpt"
        End If
    Else
        Call BonusWrk
        If chkCommPlan = 0 Then
            strRName$ = glbIHRREPORTS & "sn2291-3.rpt"
        Else
            strRName$ = glbIHRREPORTS & "sn2291-4.rpt"
        End If
    End If
End If
If glbCompSerial = "S/N - 2325W" Then
    If chkWithAct = 0 Then
        If chkCommPlan = 0 Then
            strRName$ = glbIHRREPORTS & "sn2325-1.rpt"
        Else
            strRName$ = glbIHRREPORTS & "sn2325-2.rpt"
        End If
    Else
        Call BonusWrk
        If chkCommPlan = 0 Then
            strRName$ = glbIHRREPORTS & "sn2325-3.rpt"
        Else
            strRName$ = glbIHRREPORTS & "sn2325-4.rpt"
        End If
    End If
End If
Me.vbxCrystal.ReportFileName = strRName$

x% = Cri_Sorts()   ' returns number of sections formated

If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
  
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    For x% = 0 To 3
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
    If glbCompSerial = "S/N - 2291W" Then
        Me.vbxCrystal.DataFiles(4) = Replace(UCase(glbIHRDB), "IHR001.MDB", "SN2291.MDB")
    End If
    If glbCompSerial = "S/N - 2325W" Then
        Me.vbxCrystal.DataFiles(4) = Replace(UCase(glbIHRDB), "IHR001.MDB", "SN2325.MDB")
    End If
    
    For x% = 5 To 6
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
End If

Me.vbxCrystal.WindowTitle = Me.Caption
  

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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%


Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

If comGroup(0) <> "(none)" Then
  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
Else
  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
End If
Me.vbxCrystal.GroupCondition(0) = "GROUP1;" & grpField$ & ";ANYCHANGE;A"

If GrpIdx% > 1 Then
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;X;F;X;X;X;X;X"
Else
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(1) = "GF1;T;X;X;X;X;X;X"
End If


Cri_Sorts = z% ' next section number to format

End Function



Private Function CriCheck()
Dim x%

CriCheck = False

If Len(clpDiv) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If


CriCheck = True
End Function

Private Sub Form_Load()
Screen.MousePointer = HOURGLASS
glbOnTop = "FRMRCOMPLAN"
Call comGrpLoad
Call setRptCaption(Me)
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



Private Sub BonusWrk()
Dim SQLQ, SQLX
Dim cnCP As New ADODB.Connection
Dim rsOE As New ADODB.Recordset
Dim rsCP As New ADODB.Recordset
Dim xBounsPaid
If glbCompSerial = "S/N - 2291W" Then
    cnCP.Open Replace(UCase(glbAdoIHRDB), "IHR001.MDB", "SN2291.MDB")
End If
If glbCompSerial = "S/N - 2325W" Then
    cnCP.Open Replace(UCase(glbAdoIHRDB), "IHR001.MDB", "SN2325.MDB")
End If
SQLQ = "SELECT SUM(ACT_DOLLAR) as BONUSPAID FROM HREARN "
SQLQ = SQLQ & " WHERE (LEFT(EARN_TYPE,3)='COM' OR LEFT(EARN_TYPE,3)='PER') "
rsCP.Open "COMPENSATION_PLAN", cnCP, adOpenStatic, adLockOptimistic
Do Until rsCP.EOF
    SQLX = SQLQ & " AND EMPNBR=" & rsCP("CP_EMPNBR")
    SQLX = SQLX & " AND FDATE>=" & Date_SQL(rsCP("CP_FDATE"))
    SQLX = SQLX & " AND TDATE<=" & Date_SQL(rsCP("CP_TDATE"))
    rsOE.Open SQLX, gdbAdoIhr001, adOpenStatic
    rsCP("CP_BONUSPAID") = rsOE("BONUSPAID")
    rsCP.Update
    rsOE.Close
    rsCP.MoveNext
Loop
End Sub




