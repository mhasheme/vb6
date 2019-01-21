VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRGPPosting 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Great Plains Salary Posting Report"
   ClientHeight    =   5865
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   9840
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
   ScaleHeight     =   5865
   ScaleWidth      =   9840
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkUnPosted 
      Caption         =   "Unposted"
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
      Left            =   3360
      TabIndex        =   29
      Top             =   3720
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox chkPosted 
      Caption         =   "Posted"
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
      Left            =   5520
      TabIndex        =   28
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1860
      MaxLength       =   2
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3660
      Visible         =   0   'False
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
      Index           =   1
      Left            =   1820
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "Second level of grouping records"
      Top             =   4860
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
      Left            =   1820
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Tag             =   "First Level of grouping records"
      Top             =   4500
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1680
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
      Left            =   1560
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2010
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
      Left            =   1560
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1350
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
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1020
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1560
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
      Left            =   1560
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
      Left            =   1560
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3330
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2670
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2340
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7200
      Top             =   5280
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
      Left            =   120
      TabIndex        =   27
      Top             =   3660
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
      Left            =   120
      TabIndex        =   26
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   25
      Top             =   3360
      Width           =   540
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
      Left            =   120
      TabIndex        =   24
      Top             =   2700
      Width           =   510
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   120
      TabIndex        =   23
      Top             =   3030
      Width           =   1125
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
      Left            =   120
      TabIndex        =   22
      Top             =   1050
      Width           =   615
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
      TabIndex        =   21
      Top             =   4275
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
      Left            =   120
      TabIndex        =   20
      Top             =   4920
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
      Left            =   120
      TabIndex        =   19
      Top             =   4560
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
      Left            =   0
      TabIndex        =   18
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
      Left            =   120
      TabIndex        =   17
      Top             =   2370
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
      Left            =   120
      TabIndex        =   16
      Top             =   1710
      Width           =   450
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
      Left            =   120
      TabIndex        =   15
      Top             =   1380
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
      Left            =   120
      TabIndex        =   14
      Top             =   720
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
      Left            =   120
      TabIndex        =   13
      Top             =   360
      Width           =   555
   End
End
Attribute VB_Name = "frmRGPPosting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr
 
If CriCheck() Then
    If Not PrtForm("Great Plains Salary Posting Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
    
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
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

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    
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
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub comEmpType_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkPosted_Click()
    If chkUnPosted.Value = 0 Then chkUnPosted.Value = 1
End Sub

Private Sub chkUnPosted_Click()
    If chkPosted.Value = 0 Then chkUnPosted.Value = 1
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem "(none)"
    comGroup(1).AddItem "Employee Name"
    If Not glbSQL Then
        comGroup(1).AddItem "Date of Hire"
    End If
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
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
    End Select
    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
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

Private Function Cri_SetAll()
Dim x%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere

For x% = 0 To 5
    Call Cri_Code(x%)
Next x%
Call Cri_PT
Call Cri_Shift
Call Cri_EE
Call Cri_Posted
' report name
    
If comGroup(0) <> "(none)" Then
    strRName$ = glbIHRREPORTS & "rztGPPosting.rpt"
Else
    strRName$ = glbIHRREPORTS & "rztGPPosting1.rpt"
End If
  
Me.vbxCrystal.ReportFileName = strRName$
' set to sorting/grouping criteria
x% = Cri_Sorts()   ' returns number of sections formated

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    If comGroup(0) <> "(none)" Then
        For x% = 0 To 7
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
    Else
        For x% = 0 To 3
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next x%
    End If
    ' set security for database
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
End If

' window title if appropriate
Me.vbxCrystal.WindowTitle = "Great Plains Salary Posting Report"

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

Private Sub Cri_Posted()
Dim EECri As String, OneSet%, x%

If Not (chkPosted.Value <> 0 And chkUnPosted.Value <> 0) Then
    If chkPosted.Value <> 0 Then
        EECri = "{HR_GP_CPY10140.ProcessFlag}"
    Else
        EECri = "not {HR_GP_CPY10140.ProcessFlag}"
    End If
    If chkUnPosted.Value <> 0 Then
        EECri = "not {HR_GP_CPY10140.ProcessFlag}"
    Else
        EECri = "{HR_GP_CPY10140.ProcessFlag}"
    End If
    If glbiOneWhere Then
        glbstrSelCri = glbstrSelCri & " AND " & EECri
    Else
        glbstrSelCri = EECri
    End If
End If
glbiOneWhere = True
End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

If comGroup(0) = "(none)" Then Exit Function
    
Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

GrpIdx% = comGroup(1).ListIndex
grpField$ = "{@EFullName}"
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$
      
Cri_Sorts = z% ' next section number to format

End Function

Private Function Cri_SortsDepend()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_SortsDepend = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)
If comGroup(0) = "(none)" Then Exit Function
    
Y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(x%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

Cri_SortsDepend = z% ' next section number to format

End Function

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
glbOnTop = "FRMRGPPOSTING"

Screen.MousePointer = HOURGLASS
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
Call comGrpLoad
Call setRptCaption(Me)
frmRLOA.Caption = "Great Plains Salary Posting Report"
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
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

