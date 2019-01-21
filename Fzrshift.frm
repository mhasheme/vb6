VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRShift 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Shift Schedule Report Criteria"
   ClientHeight    =   6795
   ClientLeft      =   690
   ClientTop       =   630
   ClientWidth     =   9600
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
   ScaleHeight     =   6795
   ScaleWidth      =   9600
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboPage 
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
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   5640
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
      Left            =   2010
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Tag             =   "First Level of grouping records"
      Top             =   5250
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1770
      TabIndex        =   3
      Tag             =   "EDPT-Category"
      Top             =   1320
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
      Index           =   0
      Left            =   1770
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   970
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1770
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   620
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
      Left            =   1770
      TabIndex        =   0
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
      Left            =   1770
      TabIndex        =   9
      Tag             =   "00-Enter Administered By Code"
      Top             =   3070
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1770
      TabIndex        =   10
      Tag             =   "00-Enter Section Code"
      Top             =   3420
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1770
      TabIndex        =   8
      Tag             =   "00-Enter Region Code"
      Top             =   2720
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3810
      TabIndex        =   7
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2430
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1770
      TabIndex        =   6
      Tag             =   "40-Date from and including this date forward"
      Top             =   2370
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1770
      TabIndex        =   4
      Tag             =   "10-Enter Employee Number"
      Top             =   1670
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "ED_GLNO"
      Height          =   285
      Index           =   3
      Left            =   1770
      TabIndex        =   5
      Tag             =   "00-General Ledger - Code"
      Top             =   2020
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break On"
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
      Left            =   120
      TabIndex        =   26
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblGL 
      BackStyle       =   0  'Transparent
      Caption         =   "G/L"
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
      Left            =   120
      TabIndex        =   25
      Top             =   2030
      Width           =   1575
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
      TabIndex        =   24
      Top             =   1344
      Width           =   630
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
      Left            =   120
      TabIndex        =   23
      Top             =   3465
      Width           =   540
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
      Left            =   120
      TabIndex        =   22
      Top             =   3119
      Width           =   1125
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   21
      Top             =   2776
      Width           =   510
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
      TabIndex        =   20
      Top             =   1001
      Width           =   615
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
      Top             =   5280
      Width           =   885
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
      TabIndex        =   18
      Top             =   4920
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
      Left            =   -30
      TabIndex        =   17
      Top             =   30
      Width           =   1575
   End
   Begin VB.Label lblFromDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date Range"
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
      Top             =   2433
      Width           =   1545
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
      TabIndex        =   15
      Top             =   1687
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
      Left            =   120
      TabIndex        =   14
      Top             =   658
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
      Top             =   315
      Width           =   555
   End
End
Attribute VB_Name = "frmRShift"
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
    If Not PrtForm("Shift Schedule Report Criteria", Me) Then Exit Sub
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
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ATTEND", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub comGroup_GotFocus(Index As Integer)
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    'comGroup(0).AddItem lStr("G/L")
    comGroup(0).AddItem "(none)"
    
    'cboPage.AddItem lStr("Department")
    'cboPage.AddItem lStr("G/L")
    'cboPage.AddItem lStr("Location")    'Shift
    cboPage.AddItem "(none)"
    

    comGroup(0).ListIndex = 0
    cboPage.ListIndex = 0
End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    If intIdx% = 0 Then strCd$ = "HREMP.ED_LOC"
    If intIdx% = 4 Then strCd$ = "HREMP.ED_REGION"
    If intIdx% = 5 Then strCd$ = "HREMP.ED_ADMINBY"
    If intIdx% = 6 Then strCd$ = "HREMP.ED_SECTION"  'Lucy July 4, 2000
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
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

Private Sub Cri_GLNO()
Dim EECri As String

If Len(clpCode(3).Text) > 0 Then
    EECri = "{HREMP.ED_GLNO} = '" & clpCode(3).Text & "' "
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

Private Sub Cri_Dates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%
  If Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) = 0 Then Exit Sub

TempCri = "({HREMP.ED_SENDTE} "
If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
  For x% = 0 To 1
    dtYYY% = Year(dlpDateRange(x%).Text)
    dtMM% = month(dlpDateRange(x%).Text)
    dtDD% = Day(dlpDateRange(x%).Text)
    If x% = 0 Then
      TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    Else
      TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    End If
  Next x%
Else
  If Len(dlpDateRange(0).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
  End If
  If Len(dlpDateRange(1).Text) > 0 Then
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
  End If

End If

Cri_Datst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
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
Dim RPT
Dim c As Integer

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere

Call Cri_PT
Call Cri_EE
Call Cri_Code(0)
Call Cri_Code(4)
Call Cri_Code(5)
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
Call Cri_Code(6)

Call Cri_GLNO
Call Cri_Dates

' report name

' set to sorting/grouping criteria
If comGroup(0).ListIndex <> 2 Then
    x% = Cri_Sorts()   ' returns number of sections formated
End If

'set location for database tables
If Len(glbstrSelCri) >= 0 Then
    Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

If comGroup(0).ListIndex = 2 Then
    strRName$ = glbIHRREPORTS & "sn2369shift_1.rpt"
Else
    strRName$ = glbIHRREPORTS & "sn2369shift.rpt"
End If

Me.vbxCrystal.ReportFileName = strRName$
If glbOracle Or glbSQL Then
    Me.vbxCrystal.Connect = RptODBC_SQL
End If
' window title if appropriate
Me.vbxCrystal.WindowTitle = "Shift Schedule Report"


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

'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0
Select Case comGroup(0).ListIndex
'Case 1 '0
'    grpField = "{HREMP.ED_GLNO}"
'    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(0) = grpCond$
'    grpField$ = "{HREMP.ED_DEPTNO}"
'    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(1) = grpCond$
Case 0 '1
    grpField$ = "{HRDEPT.DF_NAME}"

    Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$

    'grpField$ = "{HREMP.ED_DEPTNO}"
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
'    grpField = "{HREMP.ED_GLNO}"
'    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
'    Me.vbxCrystal.GroupCondition(1) = grpCond$
Case 1
    grpField$ = "{tblLocation.TB_DESC}"

    Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = '" & comGroup(0).Text & "'"
    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = " & grpField$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
End Select

'dscGroup$ = comGroup(0).Text
'dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
'Me.vbxCrystal.Formulas(0) = dscGroup$

''Page Breaks
'Select Case cboPage.ListIndex
''Case 1 '0 'GL
''    If comGroup(0).ListIndex = 1 Then 'GL 0
''        Me.vbxCrystal.SectionFormat(0) = "GH1;X;X;T;X;X;X;X"
''        Me.vbxCrystal.SectionFormat(1) = "GH2;X;X;F;X;X;X;X"
''        Me.vbxCrystal.SectionFormat(2) = "GF1;F;X;F;X;X;X;X"
''    Else
''        Me.vbxCrystal.SectionFormat(0) = "GH1;X;X;F;X;X;X;X"
''        Me.vbxCrystal.SectionFormat(1) = "GH2;X;X;T;X;X;X;X"
''    End If
'Case 1 '2 '1 'Location/Shift
'    'If comGroup(0).ListIndex = 1 Then '0
'        Me.vbxCrystal.SectionFormat(0) = "GH1;X;X;F;X;X;X;X"
'        Me.vbxCrystal.SectionFormat(1) = "GH2;X;X;F;X;X;X;X"
'        Me.vbxCrystal.SectionFormat(2) = "GH3;X;T;F;X;X;X;X"
'    'Else
'    '    Me.vbxCrystal.SectionFormat(0) = "GH1;X;X;F;X;X;X;X"
'    '    Me.vbxCrystal.SectionFormat(1) = "GH2;X;X;F;X;X;X;X"
'    '    Me.vbxCrystal.SectionFormat(2) = "GH3;X;T;F;X;X;X;X"
'    'End If
'Case 0 '2 'Department
'    If comGroup(0).ListIndex = 1 Then '0
'        Me.vbxCrystal.SectionFormat(0) = "GH1;X;X;F;X;X;X;X"
'        Me.vbxCrystal.SectionFormat(1) = "GH2;X;X;T;X;X;X;X"
'    Else
'        Me.vbxCrystal.SectionFormat(0) = "GH1;X;X;T;X;X;X;X"
'        Me.vbxCrystal.SectionFormat(1) = "GH2;X;X;F;X;X;X;X"
'    End If
'Case 2 'None
''    If comGroup(0).ListIndex = 1 Then '0
'        Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;F;X;X;X;X"
'        Me.vbxCrystal.SectionFormat(1) = "GH2;X;F;F;X;X;X;X"
'        'Me.vbxCrystal.SectionFormat(2) = "GH3;X;F;F;X;X;X;X"
'        'Me.vbxCrystal.SectionFormat(4) = "GH4;X;F;F;X;X;X;X"
''    Else
''        Me.vbxCrystal.SectionFormat(0) = "GH1;X;F;F;X;X;X;X"
''        Me.vbxCrystal.SectionFormat(1) = "GH2;X;F;F;X;X;X;X"
''        Me.vbxCrystal.SectionFormat(2) = "GH3;X;F;F;X;X;X;X"
''        Me.vbxCrystal.SectionFormat(4) = "GH4;X;F;F;X;X;X;X"
''    End If
'Case Else
'    If comGroup(0).ListIndex = 1 Then 'GL 0
'        Me.vbxCrystal.SectionFormat(0) = "GF1;F;X;X;X;X;X;X"
'    End If
'End Select

'Call setRptLabel(Me, 1)

Cri_Sorts = z% ' next section number to format

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

'If Not clpCode(0).ListChecker Then Exit Function
'For x% = 3 To 6
'    If Not clpCode(x).ListChecker Then Exit Function
'Next x%
'

 If Len(dlpDateRange(0).Text) > 0 Then
    If Not IsDate(dlpDateRange(0).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(0).Text = ""
        dlpDateRange(0).SetFocus
        Exit Function
    End If
 End If
 If Len(dlpDateRange(1).Text) > 0 Then
    If Not IsDate(dlpDateRange(1).Text) Then
        MsgBox "Not a valid date"
        dlpDateRange(1).Text = ""
        dlpDateRange(1).SetFocus
        Exit Function
    End If
 End If
    'check to ensure that the from date is <= the to date
 If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
      If CVDate(dlpDateRange(0).Text) > CVDate(dlpDateRange(1).Text) Then
        MsgBox "Not a valid date range"
        dlpDateRange(1).Text = ""
        dlpDateRange(1).SetFocus
        Exit Function
      End If
 End If

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

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

glbOnTop = Me.name

Screen.MousePointer = HOURGLASS

Call comGrpLoad
Call setRptCaption(Me)

lblGL.Caption = lStr("G/L")

If glbLinamar Then clpCode(4).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6
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
Call SetPanHelp(Me.ActiveControl)
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
