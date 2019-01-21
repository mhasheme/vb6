VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSFRXMLTable 
   Caption         =   "XML Working Table Report"
   ClientHeight    =   6690
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6690
   ScaleWidth      =   9720
   WindowState     =   2  'Maximized
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "First Level of grouping records"
      Top             =   2730
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2115
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "Year"
      Top             =   1665
      Visible         =   0   'False
      Width           =   900
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Tag             =   "00-Enter Union Code"
      Top             =   2040
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7200
      Top             =   6000
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
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   0
      Left            =   1725
      TabIndex        =   8
      Tag             =   "41-Effective date "
      Top             =   840
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1300
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      Height          =   285
      Index           =   1
      Left            =   3840
      TabIndex        =   9
      Tag             =   "41-Effective date "
      Top             =   840
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1300
   End
   Begin VB.Label lblDet 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label lblDet 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   840
      Width           =   195
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2010
      Visible         =   0   'False
      Width           =   420
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
      TabIndex        =   6
      Top             =   360
      Width           =   1575
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
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   360
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblYear 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   3
      Top             =   1665
      Visible         =   0   'False
      Width           =   330
   End
End
Attribute VB_Name = "frmSFRXMLTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "frmSFRXMLTable"

Screen.MousePointer = HOURGLASS

Call comGrpLoad

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem ("Union")
    comGroup(0).AddItem ("Year")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 2
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

Private Sub txtYear_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    Screen.MousePointer = HOURGLASS
    x% = Cri_SetAll()
    Me.vbxCrystal.WindowTitle = glbFrmCaption$
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
Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then

    If glbFrmCaption$ = "PA Details Report" Then
        If Not PrtForm("PA Details Report Criteria", Me) Then Exit Sub
    End If
    If glbFrmCaption$ = "PA Master Report" Then
        If Not PrtForm("PA Master Report Criteria", Me) Then Exit Sub
    End If

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
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub
Private Function CriCheck()
Dim x%

CriCheck = False
    
    If Len(dlpDate(0).Text) > 0 Then
        If Not IsDate(dlpDate(0).Text) Then
            MsgBox "Invalid From Date."
            dlpDate(0).SetFocus
            Exit Function
        End If
    End If
    If Len(dlpDate(1).Text) > 0 Then
        If Not IsDate(dlpDate(1).Text) Then
            MsgBox "Invalid To Date"
            dlpDate(1).SetFocus
            Exit Function
        End If
    End If
    
'If Len((txtYear.Text)) > 0 Then
'    If Not IsNumeric(txtYear.Text) Then
'        MsgBox "Invalid Year"
'        txtYear.SetFocus
'        Exit Function
'    End If
'End If
'
'For X% = 0 To 0
'    If Not clpCode(X).ListChecker Then Exit Function
'Next X%


CriCheck = True
End Function


Private Function Cri_SetAll()
Dim x%

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call Cri_FTDates

'For X% = 0 To 0
'    Call Cri_Code(X%)
'Next X%
'
'Call Cri_Year

strRName$ = glbIHRREPORTS & "RSF_XMLImp.rpt"
Me.vbxCrystal.ReportFileName = strRName$
'' set to sorting/grouping criteria
'X% = Cri_Sorts()   ' returns number of sections formated
'
'
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If
  

Me.vbxCrystal.Connect = RptODBC_SQL

  

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


Private Sub Cri_Year()
Dim YearCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(txtYear.Text) > 0 Then
    YearCri = "({@xYear} = " & txtYear.Text & ")"
End If

If Len(YearCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = YearCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & YearCri
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
    Case 0: strCd$ = "HRP_HOURLY_PEN_RATES.PE_UNION"
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

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
y% = 0
grpField$ = getEGroup(comGroup(0).Text)

'Call setRptLabel(Me, 0)
If comGroup(0) = "(none)" Then
    grpField$ = "{HRP_HOURLY_PEN_RATES.PE_COMPNO}"
End If
If comGroup(0) = "Union" Then
    grpField$ = "{HRP_HOURLY_PEN_RATES.PE_UNION}"
End If
If comGroup(0) = "Year" Then
    grpField$ = "{@xYear}"
End If
y% = x% + 1
dscGroup$ = comGroup(x%).Text
dscGroup$ = "descGroup" & CStr(y%) & "= '" & dscGroup$ & "'"
'Me.vbxCrystal.Formulas(X%) = dscGroup$

grpCond$ = "GROUP" & CStr(y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(x%) = grpCond$

strSFormat$ = "GH1;T;T;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1
strSFormat$ = "GF1;T;X;X;X;X;X;X"
Me.vbxCrystal.SectionFormat(z%) = strSFormat$
z% = z% + 1

''GrpIdx% = comGroup(1).ListIndex
''Select Case GrpIdx%
''    Case 0: grpField$ = "{@EFullName}"
''End Select
''grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
''Me.vbxCrystal.GroupCondition(1) = grpCond$

Cri_Sorts = z% ' next section number to format

End Function


Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpDate(0).Text) > 0 And Len(dlpDate(1).Text) > 0 Then
    TempCri = "({HRSF_XML_IMPORT.SF_FILEDATE} "
    dtYYY% = Year(dlpDate(0).Text)
    dtMM% = Month(dlpDate(0).Text)
    dtDD% = Day(dlpDate(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDate(1).Text)
    dtMM% = Month(dlpDate(1).Text)
    dtDD% = Day(dlpDate(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDate(x%).Text) > 0 Then
        TempCri = "({HRSF_XML_IMPORT.SF_FILEDATE} "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDate(x%).Text)
        dtMM% = Month(dlpDate(x%).Text)
        dtDD% = Day(dlpDate(x%).Text)
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
