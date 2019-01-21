VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRCompCost 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Company Associated Cost Report"
   ClientHeight    =   5640
   ClientLeft      =   390
   ClientTop       =   1185
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
   ScaleHeight     =   5640
   ScaleWidth      =   9840
   WindowState     =   2  'Maximized
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
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Tag             =   "First Level of grouping records"
      Top             =   4050
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
      Index           =   1
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Second level of grouping records"
      Top             =   4365
      Width           =   2325
   End
   Begin VB.Frame fraDetail 
      Caption         =   "Report Type"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   3255
      Begin VB.OptionButton optSummary 
         Caption         =   "Summary"
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
         Left            =   1920
         TabIndex        =   5
         Tag             =   "Summary Report"
         Top             =   270
         Width           =   1215
      End
      Begin VB.OptionButton optDetails 
         Caption         =   "Detail"
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
         Left            =   240
         TabIndex        =   4
         Tag             =   "Detail Report"
         Top             =   270
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtFirmNo 
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
      Left            =   2230
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "00-Enter Firm Number"
      Top             =   870
      Width           =   1455
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6600
      Top             =   3720
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
   Begin INFOHR_Controls.DateLookup dlpStateDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "40-Statement Date"
      Top             =   1800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3870
      TabIndex        =   2
      Tag             =   "40-Date upto and including this date forward"
      Top             =   1320
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   1
      Tag             =   "40-Date from and including this date forward"
      Top             =   1320
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Date"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   1110
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " - "
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
      Left            =   3600
      TabIndex        =   15
      Top             =   1365
      Width           =   135
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
      TabIndex        =   14
      Top             =   4080
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
      Left            =   330
      TabIndex        =   13
      Top             =   4395
      Width           =   660
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
      TabIndex        =   12
      Top             =   3840
      Width           =   1575
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
      TabIndex        =   10
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date Range"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1335
      Width           =   1260
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Firm #"
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
      Left            =   240
      TabIndex        =   8
      Top             =   870
      Width           =   435
   End
End
Attribute VB_Name = "frmRCompCost"
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
    If frmRCompCost.Caption = "Company Associated Cost Report" Then
        If Not PrtForm("Company Associated Cost Report Criteria", Me) Then Exit Sub
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
MsgBox "Error Printing - check your Windows Printer setup"
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
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
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
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "HR_OHS_COMPANY_COST", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub Cri_Code(intIdx%)
'Dim CodeCri As String
'Dim strCd$
'Dim tbl As String
'
'If Len(clpCode(intIdx%).Text) > 0 Then
'
'    Select Case intIdx%
'    Case 0: strCd$ = "HREMP.ED_LOC"
'    Case 3: strCd$ = "HREMP.ED_REGION"
'    Case 4: strCd$ = "HREMP.ED_ADMINBY"
'    Case 5: strCd$ = "HREMP.ED_SECTION"
'    Case 1: strCd$ = "HREMP.ED_EMP"
'    Case 2: strCd$ = "HREMP.ED_ORG"
'    Case 6: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_WCBRES"
'    Case Else: strCd$ = "HR_OCC_HEALTH_SAFETY.EC_TYPE"
'    End Select
'        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
'    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
'        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
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
'
'
'End Sub
'
'Private Sub Cri_Div()
'
'Dim DivCri As String
'Dim countr   As Integer  ' EEList_Snap is definded at form level
'
'
'If Len(clpDiv.Text) > 0 Then
'        DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
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
'
'Private Sub Cri_DOW()
'Dim EECri As String, OneSet%
'Dim x%
'
'
'For x% = 0 To 1
'    If Len(medDOW(x%)) > 0 Then
'        OneSet% = OneSet% + 1
'    End If
'Next x%
'
'If OneSet% = 0 Then Exit Sub
'
'
'For x% = 0 To 1
'    If Len(medDOW(x%)) > 0 Then
'        EECri = EECri & medDOW(x%)
'        OneSet% = OneSet% - 1
'    End If
'    If OneSet% > 0 Then
'        EECri = EECri & " ,"
'    End If
'Next x%
'EECri = "{@DOW} in [" & EECri & "]"
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True
'
'
'End Sub
'
'Private Sub Cri_EE()
'Dim EECri As String
'If Len(elpEEID.Text) <= 0 Then Exit Sub
'
'If Len(elpEEID.Text) > 0 Then
'        EECri = "{HREMP.ED_EMPNBR} IN [" & getEmpnbr(elpEEID.Text) & "] "
'End If
'
'
'If Len(EECri) >= 1 Then
'    If glbiOneWhere Then
'        glbstrSelCri = glbstrSelCri & " AND " & EECri
'    Else
'        glbstrSelCri = EECri
'    End If
'    glbiOneWhere = True
'End If
'
'
'End Sub

Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%
Dim tbl As String

If Len(dlpDateRange(0).Text) <= 0 And Len(dlpDateRange(1).Text) <= 0 Then Exit Sub

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({HR_OHS_COMPANY_COST.CA_FDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & "in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")"

    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "

    TempCri = TempCri & " AND ({HR_OHS_COMPANY_COST.CA_FDATE} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & "in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")"
    
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    
    GoTo Cri_FTDatst
End If

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

Private Sub Cri_StatementDate()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%
Dim tbl As String

If Len(dlpStateDate.Text) <= 0 Then Exit Sub

If Len(dlpStateDate.Text) > 0 Then
    TempCri = "({HR_OHS_COMPANY_COST.CA_STMTDT} "
    dtYYY% = Year(dlpStateDate.Text)
    dtMM% = month(dlpStateDate.Text)
    dtDD% = Day(dlpStateDate.Text)
    TempCri = TempCri & " = Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & "))"
End If

If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
    End If
    glbiOneWhere = True
End If

End Sub

'Private Sub Cri_Shift()
'Dim EECri As String, OneSet%, x%
'
'If Len(txtShift.Text) < 1 Then Exit Sub
'
'    EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True
'End Sub
'
'Private Sub Cri_PT()
'Dim EECri As String, OneSet%, x%
'
'If Len(clpPT.Text) < 1 Then Exit Sub
'
'EECri = "{HREMP.ED_PT}= '" & clpPT.Text & "'"
'
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True
'
'
'End Sub
'
'Private Sub Cri_Hazard()
'Dim BodyCri As String
'
'If cmbHRat.Text <> "" Then
'    BodyCri = "{HR_OCC_HEALTH_SAFETY.EC_HAZARD} = '" & cmbHRat.Text & "'"
'Else
'    Exit Sub
'End If
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & BodyCri
'Else
'    glbstrSelCri = BodyCri
'End If
'glbiOneWhere = True
'
'End Sub


Private Function Cri_SetAll()
Dim x%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
'Call glbCri_DeptUN(clpDept.Text)

'Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
'For x% = 0 To 6
'    Call Cri_Code(x%)
'Next
'Call Cri_EE       'laura dec 03, 1997
'Call Cri_PT
'Call Cri_Shift
'Call Cri_Hazard
Call Cri_FirmNo
Call Cri_FTDates
Call Cri_StatementDate


If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

x% = Cri_Sorts

' report name
If optDetails.Value = True Then
    If comGroup(0) <> "(none)" Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzCompCstD.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzCompCstD1.rpt"
    End If
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
        
    '    For x% = 0 To 1
    '        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    '    Next x%
    'End If

    Me.vbxCrystal.WindowTitle = "Company Associated Cost Report"

ElseIf optSummary.Value = True Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzCompCstS.rpt"
    'If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = "PWD=petman;"
        
    '    For x% = 0 To 1
    '        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    '    Next x%
    'End If
    Me.vbxCrystal.WindowTitle = "Company Associated Cost Summary Report"
End If

Cri_SetAll = True

Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "FR Comp Cost", "Comp Cost Report", "Select")
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

'' first set primary grouping ------------------------- **
Y% = 0
'grpField$ = getEGroup(comGroup(0).Text)
'If comGroup(0) = "Benefit Group" Then grpField = "{tblBGroup.TB_DESC}"
If comGroup(0) = "(none)" Then Exit Function
GrpIdx% = comGroup(0).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{HR_OHS_COMPANY_COST.CA_FIRMNO}"
    Case 1: grpField$ = "{HR_OHS_COMPANY_COST.CA_FDATE}"
End Select
    
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
    Case 0: grpField$ = "{HR_OHS_COMPANY_COST.CA_FIRMNO}"
    Case 1: grpField$ = "{HR_OHS_COMPANY_COST.CA_FDATE}"
End Select
grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

Cri_Sorts = z% ' next section number to format

End Function

Private Sub Cri_FirmNo()
Dim EECri As String, OneSet%, x%

If Len(txtFirmNo.Text) < 1 Then Exit Sub


EECri = "{HR_OHS_COMPANY_COST.CA_FIRMNO}= '" & txtFirmNo.Text & "'"


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True

End Sub

Private Function CriCheck()
Dim x%

CriCheck = False

'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
'    MsgBox lStr("If Division Entered - it must be known")
'     clpDiv.SetFocus
'    Exit Function
'End If
'
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
'    MsgBox "If Department Entered - it must be known"
'     clpDept.SetFocus
'    Exit Function
'End If
'
'For x% = 0 To 5
'    If Not clpCode(x).ListChecker Then Exit Function
'Next x%
'
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
'    MsgBox lStr("Category code must be valid")
'     clpPT.SetFocus
'    Exit Function
'End If

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

'If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text = "" Then
'  MsgBox "If entering a date, both from and to date must be entered!"
'  dlpDateRange(0).Text = ""
'  dlpDateRange(0).SetFocus
'  Exit Function
'End If
'
'If dlpDateRange(1).Text <> "" And dlpDateRange(0).Text = "" Then
'  MsgBox "If entering a date, both from and to date must be entered!"
'  dlpDateRange(1).Text = ""
'  dlpDateRange(1).SetFocus
'  Exit Function
'End If

If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
  If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
    MsgBox "From Date must be less than To Date!"
    dlpDateRange(1).SetFocus
    Exit Function
  End If
End If

'For X% = 0 To 1
' If Len(medDOW(X%)) > 0 Then
'    If medDOW(X%) < 0 Or medDOW(X%) > 7 Then
'        MsgBox "Day of Week must be between 1 and 7"
'        medDOW(X%).SetFocus
'        Exit Function
'    End If
'  End If
'Next X%
'If Not elpEEID.ListChecker Then
'    Exit Function
'End If
 

CriCheck = True
End Function

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpDateRange_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub dlpStateDate_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Sub comGrpLoad()
Dim x%

    comGroup(0).Clear
    comGroup(1).Clear

    comGroup(0).AddItem lStr("Firm Number")
    comGroup(0).AddItem "From Date"
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    
    comGroup(1).AddItem lStr("Firm Number")
    comGroup(1).AddItem "From Date"
    comGroup(1).AddItem "(none)"
    comGroup(1).ListIndex = 0
    
End Sub

'Private Sub clpCode_GotFocus(Index As Integer)
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmbHRat_Change()
''Ticket # 6831 - For Burlington Tech.
'If glbCompSerial = "S/N - 2351W" Then
'    If cmbHRat = "" Then
'        txtHRat = ""
'    End If
'End If
'End Sub

'Private Sub cmbHRat_Click()
''Ticket # 6831 - For Burlington Tech.
'If glbCompSerial = "S/N - 2351W" Then
'    Select Case cmbHRat
'        Case "Frequent": txtHRat = "F"
'        Case "Occassional": txtHRat = "O"
'        Case "Rare": txtHRat = "R"
'        Case Else
'            txtHRat = ""
'    End Select
'Else
'    txtHRat = cmbHRat
'End If
'End Sub
'
'Private Sub cmbHRat_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

Call comGrpLoad

'If Not glbMulti Then
'    lblShift.Visible = True
'    txtShift.Visible = True
'End If
'Call setRptCaption(Me)
'If glbLinamar Then clpCode(3).MaxLength = 8
'If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
'If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

''Ticket # 6831 - For Burlington Tech.
'If glbCompSerial = "S/N - 2351W" Then
'    lblTitle(9).Caption = "Frequency"
'    cmbHRat.Width = 1500
'    'cmbHRat.Left = 7560
'    cmbHRat.AddItem ""
'    cmbHRat.AddItem "Frequent"
'    cmbHRat.AddItem "Occassional"
'    cmbHRat.AddItem "Rare"
'Else
'    cmbHRat.AddItem ""
'    cmbHRat.AddItem "A"
'    cmbHRat.AddItem "B"
'    cmbHRat.AddItem "C"
'End If

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

Private Sub optDetails_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optSummary_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub txtShift_GotFocus()
'Call SetPanHelp(Me.ActiveControl)
'End Sub

Private Sub txtFirmNo_GotFocus()
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

'Private Sub SELATTWRK()
''Created by Bryan 6/Jul/05 Ticket #8857
'Dim xlen, xxx, xx1
'Dim db001 As Database
'Dim SQLQ
'Dim xFieldList
'
'On Error GoTo AttWrkError
'MDIMain.panHelp(0).FloodType = 1
'MDIMain.panHelp(1).Caption = " Please Wait"
'MDIMain.panHelp(2).Caption = ""
'MDIMain.panHelp(0).FloodPercent = 1
'MDIMain.panHelp(0).FloodPercent = 3
'gdbAdoIhr001.CommandTimeout = 300
'Screen.MousePointer = HOURGLASS
'MDIMain.panHelp(0).FloodPercent = 15
'gdbAdoIhr001.BeginTrans
'SQLQ = "DELETE FROM HRWSIBWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
'gdbAdoIhr001.Execute SQLQ
'gdbAdoIhr001.CommitTrans
'If Not glbSQL And Not glbOracle Then Call Pause(1)
'MDIMain.panHelp(0).FloodPercent = 30
'
'xFieldList = Get_Fields(gdbAdoIhr001W, "HRWSIBWRK", "WRKEMP")
'SQLQ = "INSERT INTO HRWSIBWRK (" & xFieldList & ",WRKEMP) "
'SQLQ = SQLQ & in_SQL(glbIHRDBW)
'SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP "
'If Not glbOracle Then
'    'Ticket #16854
'    'SQLQ = SQLQ & " FROM Term_HROHSCOS LEFT JOIN Term_HREMP ON Term_HROHSCOS.CC_EMPNBR = Term_HREMP.ED_EMPNBR "
'    SQLQ = SQLQ & " FROM Term_HROHSCOS LEFT JOIN Term_HREMP ON Term_HROHSCOS.TERM_SEQ = Term_HREMP.TERM_SEQ "
'Else
'    'Ticket #16854
'    'SQLQ = SQLQ & " FROM TERM_HROHSCOS, TERM_HREMP WHERE TERM_HROHSCOS.CC_EMPNBR = TERM_HREMP.ED_EMPNBR(+) "
'    SQLQ = SQLQ & " FROM TERM_HROHSCOS, TERM_HREMP WHERE TERM_HROHSCOS.TERM_SEQ = TERM_HREMP.TERM_SEQ(+) "
'End If
'
'
'MDIMain.panHelp(0).FloodPercent = 45
'gdbAdoIhr001.BeginTrans
'gdbAdoIhr001.Execute SQLQ, , adCmdText
'gdbAdoIhr001.CommitTrans
'If Not glbSQL And Not glbOracle Then Call Pause(1)
'
'MDIMain.panHelp(0).FloodPercent = 50
'SQLQ = "INSERT INTO HRWSIBWRK (" & xFieldList & ",WRKEMP) "
'SQLQ = SQLQ & in_SQL(glbIHRDBW)
'SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP "
'If Not glbOracle Then
'    SQLQ = SQLQ & " FROM Term_HROHSCOS LEFT JOIN Term_HREMP ON Term_HROHSCOS.CC_EMPNBR = Term_HREMP.ED_EMPNBR "
'Else
'    SQLQ = SQLQ & " FROM TERM_HROHSCOS, TERM_HREMP WHERE TERM_HROHSCOS.CC_EMPNBR = TERM_HREMP.ED_EMPNBR(+) "
'End If
'SQLQ = SQLQ & "WHERE CC_WCBNBR IS NOT NULL"
'
'
'MDIMain.panHelp(0).FloodPercent = 60
'gdbAdoIhr001X.BeginTrans
'gdbAdoIhr001X.Execute SQLQ
'gdbAdoIhr001X.CommitTrans
'If Not glbSQL And Not glbOracle Then Call Pause(1)
''Commented by Bryan 03/Oct/05 Ticket#9419
''xFieldList = Get_Fields(gdbAdoIhr001W, "HROCCHSWRK", "WRKEMP")
''
''SQLQ = "INSERT INTO HROCCHSWRK (" & xFieldList & ",WRKEMP) "
''SQLQ = SQLQ & in_SQL(glbIHRDBW)
''SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP FROM Term_HR_OCC_HEALTH_SAFETY"
''
''MDIMain.panHelp(0).FloodPercent = 75
''gdbAdoIhr001X.BeginTrans
''gdbAdoIhr001X.Execute SQLQ
''gdbAdoIhr001X.CommitTrans
''If Not glbSQL And Not glbOracle Then Call Pause(1)
''
''SQLQ = "INSERT INTO HROCCHSWRK (" & xFieldList & ",WRKEMP) "
''SQLQ = SQLQ & in_SQL(glbIHRDBW)
''SQLQ = SQLQ & " SELECT " & xFieldList & ",'" & glbUserID & "' AS WRKEMP FROM HR_OCC_HEALTH_SAFETY"
''
''MDIMain.panHelp(0).FloodPercent = 80
''gdbAdoIhr001.BeginTrans
''gdbAdoIhr001.Execute SQLQ
''gdbAdoIhr001.CommitTrans
''If Not glbSQL And Not glbOracle Then Call Pause(1)
''
'gdbAdoIhr001.CommandTimeout = 300
'MDIMain.panHelp(0).FloodPercent = 100
'MDIMain.panHelp(0).FloodType = 0
'MDIMain.panHelp(1).Caption = ""
'MDIMain.panHelp(2).Caption = ""
'Exit Sub
'
'AttWrkError:
'    gdbAdoIhr001.CommandTimeout = 300
'    MDIMain.panHelp(0).FloodType = 0
'    MDIMain.panHelp(1).Caption = ""
'    MDIMain.panHelp(2).Caption = ""
'    'If Err.Number = 2147217871 Then MsgBox Err.Description
'    ' dkostka - 04/18/01 - Not sure why the previous line was hiding errors,
'    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
'    Exit Sub
'
'End Sub
'
'
