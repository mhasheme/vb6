VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmREmpEquity 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Employment Equity Report"
   ClientHeight    =   7410
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   10380
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7410
   ScaleWidth      =   10380
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtProv 
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
      Left            =   2280
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5280
      Width           =   375
   End
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "FZEmpEquity.frx":0000
      Left            =   2280
      List            =   "FZEmpEquity.frx":0002
      TabIndex        =   13
      Tag             =   "Type: Active or Terminated "
      Top             =   4560
      Width           =   1695
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      MaxLength       =   4
      TabIndex        =   12
      Tag             =   "00-Shift"
      Top             =   4220
      Visible         =   0   'False
      Width           =   450
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   0
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
      PrintFileLinesPerPage=   60
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3570
      TabIndex        =   8
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2860
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   1950
      TabIndex        =   11
      Tag             =   "EDSE-Section "
      Top             =   3880
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1950
      TabIndex        =   10
      Tag             =   "EDAB-Administered By"
      Top             =   3540
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1950
      TabIndex        =   9
      Tag             =   "EDRG-Region"
      Top             =   3200
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1950
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2180
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
      Index           =   3
      Left            =   1950
      TabIndex        =   4
      Top             =   1840
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
      Index           =   0
      Left            =   1950
      TabIndex        =   2
      Tag             =   "EDLC-Location"
      Top             =   1160
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1950
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   820
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
      Left            =   1950
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
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1950
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2520
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1950
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1500
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   1950
      TabIndex        =   7
      Tag             =   "40-Date from and including this date forward"
      Top             =   2860
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpPlanNbr 
      Height          =   285
      Left            =   1950
      TabIndex        =   14
      Tag             =   "11-Enter Plan Number"
      Top             =   4930
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   10
      LookupType      =   7
   End
   Begin INFOHR_Controls.CodeLookup clpProv 
      Height          =   285
      Left            =   2790
      TabIndex        =   16
      Tag             =   "31-Province - Code"
      Top             =   5280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Province"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   150
      TabIndex        =   32
      ToolTipText     =   "Compares with Employment Equity fields"
      Top             =   5325
      Width           =   630
   End
   Begin VB.Label lblPlanNbr 
      Appearance      =   0  'Flat
      Caption         =   "Plan Number"
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
      Left            =   150
      TabIndex        =   31
      Top             =   4971
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Index           =   0
      Left            =   150
      TabIndex        =   30
      Top             =   4629
      Width           =   480
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From / To Date"
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
      Left            =   150
      TabIndex        =   29
      Top             =   2919
      Width           =   1095
   End
   Begin VB.Label lblShift 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Left            =   150
      TabIndex        =   28
      Top             =   4287
      Visible         =   0   'False
      Width           =   315
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
      Left            =   150
      TabIndex        =   27
      Top             =   1551
      Width           =   420
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
      Left            =   150
      TabIndex        =   26
      Top             =   1893
      Width           =   450
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
      Left            =   150
      TabIndex        =   25
      Top             =   1209
      Width           =   615
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
      Left            =   150
      TabIndex        =   24
      Top             =   3603
      Width           =   1125
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   23
      ToolTipText     =   "Compares with Employment Equity fields"
      Top             =   2235
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
      Left            =   150
      TabIndex        =   22
      Top             =   3945
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
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      TabIndex        =   21
      ToolTipText     =   "Compares with Employment Equity fields"
      Top             =   3261
      Width           =   510
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
      TabIndex        =   20
      Top             =   150
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
      Left            =   150
      TabIndex        =   19
      Top             =   2577
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
      Left            =   150
      TabIndex        =   18
      Top             =   867
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
      Left            =   150
      TabIndex        =   17
      Top             =   525
      Width           =   555
   End
End
Attribute VB_Name = "frmREmpEquity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fsnapEENames As Recordset
Dim DATE1, DATE2

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%
On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm(Me.Caption, Me) Then Exit Sub
    Call set_PrintState(False)
    X% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
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
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Call set_PrintState(False)
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    Screen.MousePointer = HOURGLASS
    X% = Cri_SetAll()
    'Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Screen.MousePointer = DEFAULT
    'Me.vbxCrystal.Action = 1
    'vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
End If
Exit Sub

CRW_Err:

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Timesheet Status", "Timesheet Status Report", "Select")
Cri_SetAll = False
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
    Case 1: strCd$ = "HREMP.EQ_REGION"
    Case 2: strCd$ = "HREMP.ED_SECTION"
    Case 3: strCd$ = "HREMP.ED_EMP"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_ORG"
    Case 6: strCd$ = "HREMP.ED_PT"
    End Select
    'CodeCri = "(" & strCd$ & " = '" & clpCode(intIdx%).Text & "')"
    CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
    
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
Dim countr   As Integer  ' EEList_Snap is definded at form level


If Len(clpDiv.Text) > 0 Then
    DivCri = "(HREMP.ED_DIV in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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
    EECri = "HREMP.ED_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
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
    Dim X%, strRName$
    
    Cri_SetAll = False
    
    '''On Error GoTo modSetCriteria_Err
    Screen.MousePointer = HOURGLASS
    
    glbiOneWhere = False
    glbstrSelCri = ""
    
    Call glbCri_DeptUN(clpDept.Text)
    
    Call Cri_Div
    For X% = 0 To 6
        If X% <> 1 And X% <> 6 Then
            Call Cri_Code(X%)
        End If
    Next X%
    'Call Cri_EE
    Call Cri_Shift
    
    Call Export_Employment_Equity

    'strRName$ = glbIHRREPORTS & "rztimesheet.rpt"
    
    'Me.vbxCrystal.ReportFileName = strRName$
    ''x% = Cri_Sorts()   ' returns number of sections formated
    'Me.vbxCrystal.SelectionFormula = "{HR_ATT_TIMESHEET.AD_WRKEMP}='" & glbUserID & "'"
    'Me.vbxCrystal.WindowTitle = Me.Caption
    'If glbSQL Or glbOracle Then
    '    Me.vbxCrystal.Connect = RptODBC_SQL
    'Else
    '    Me.vbxCrystal.Connect = RptODBC_SQL
    '    Me.vbxCrystal.Connect = "PWD=petman;"
    '    Me.vbxCrystal.DataFiles(0) = glbIHRDBW
    '    'For x% = 1 To 7
    '    '    Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    '    'Next x%
    'End If
    
    Cri_SetAll = True

    Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Call Generate Report", "Employment Equity Report", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Function CriCheck()
Dim X%

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

For X% = 0 To 6
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not elpEEID.ListChecker Then
    Exit Function
End If

If clpPlanNbr.Caption = "Unassigned" Then
    MsgBox "Invalid Plan Number"
    clpPlanNbr.SetFocus
    Exit Function
End If

If clpProv.Caption = "Unassigned" Then
    MsgBox "Invalid Province"
    clpProv.SetFocus
    Exit Function
End If

If Len(dlpDateRange(0)) = 0 Or Len(dlpDateRange(1)) = 0 Then
    MsgBox "From / To Date is mandatory"
    dlpDateRange(0).SetFocus
    Exit Function
End If

If Not IsDate(dlpDateRange(0)) Then
    MsgBox "Invalid From Date"
    dlpDateRange(0).SetFocus
    Exit Function
End If
    
If Not IsDate(dlpDateRange(1)) Then
    MsgBox "Invalid To Date"
    dlpDateRange(1).SetFocus
    Exit Function
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
End If


CriCheck = True

End Function

Private Sub clpProv_Change()
    txtProv.Text = Get_ProvinceCodeData(clpProv, "NBR")
End Sub

Private Sub cmbType_Change()
    Call cmbType_Click
End Sub

Private Sub cmbType_Click()
    If cmbType.ListIndex = 0 Then
        'elpEEID.ShowDescription = True
        elpEEID.LookupType = 0
        elpEEID.Enabled = True
    ElseIf cmbType.ListIndex = 1 Then
        elpEEID.LookupType = 1
        elpEEID.Enabled = True
    Else
        elpEEID.Enabled = False
    End If
End Sub

Private Sub elpEEID_Change()
    If Len(elpEEID) > 0 Then
        Call cmbType_Click
    End If
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Screen.MousePointer = HOURGLASS
glbOnTop = Me.name

Call setCaption(lblDiv)
Call setCaption(lblRegion)
Call setCaption(lblSection)
Call setCaption(lblDept)
Call setCaption(lblEENum(1))
Call setRptCaption(Me)
'lblFromTo.Caption = lStr("From Date") & " / " & lStr("To Date")


If glbCompSerial = "S/N - 2227W" Then clpCode(1).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

cmbType.Clear
cmbType.AddItem "Active"
cmbType.AddItem "Terminated"
cmbType.AddItem "All"
cmbType.ListIndex = 2

Call INI_Controls(Me)

'chkShowEmp.Visible = True
If glbLinamar Then
    clpCode(1).MaxLength = 8
End If

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

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

Private Sub Cri_Dept()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim DeptCri As String

'Ticket #21968 - Allow multi code selection criteria
'If Len(clpDept.Text) > 0 Then DeptCri = " AND (ED_DEPTNO = '" & clpDept.Text & "') "
If Len(clpDept.Text) > 0 Then DeptCri = " AND (ED_DEPTNO in ['" & Replace(clpDept.Text, ",", "','") & "']) "

glbstrSelCri = glbSeleDeptUn & DeptCri
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

Function StripChar(StringToStrip, CharToStrip)
    Dim I, buf, OneChar
    
    For I = 1 To Len(StringToStrip)
        OneChar = Mid(StringToStrip, I, 1)
        If OneChar <> CharToStrip Then buf = buf & OneChar
    Next I
    StripChar = buf
End Function

Function getDateRange(theClientNumber, thePayNbr, theYear)
    Dim rsPayPeriod As New ADODB.Recordset
    Dim SQLQ, intNum
    
    On Error Resume Next
    
    getDateRange = "|"
    
    If Not IsNumeric(thePayNbr) Then Exit Function
    If Not IsNumeric(theYear) Then Exit Function
    
    SQLQ = "SELECT PP_NBR,PP_YEAR,PP_Start,PP_End FROM HR_PAYPERIOD WHERE PP_PAYP='" & theClientNumber & "'"
    SQLQ = SQLQ & " and PP_NBR = " & thePayNbr
    SQLQ = SQLQ & " and PP_YEAR = '" & theYear & "'"
    rsPayPeriod.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsPayPeriod.EOF Then
        getDateRange = rsPayPeriod("PP_Start") & "|" & rsPayPeriod("PP_End")
    End If
    rsPayPeriod.Close
    Exit Function

End Function

Private Function getWSQLQ(WithAtt As Boolean)
Dim QStr
QStr = glbSeleDeptUn
If clpDiv <> "" Then QStr = QStr & " AND ED_DIV in ('" & Replace(clpDiv, ",", "','") & "')"
If clpCode(0) <> "" Then QStr = QStr & " AND ED_LOC='" & clpCode(0) & "'"
If clpCode(1) <> "" Then QStr = QStr & " AND ED_ORG in ('" & Replace(clpCode(1), ",", "','") & "')"
If clpCode(2) <> "" Then QStr = QStr & " AND ED_EMP in ('" & Replace(clpCode(2), ",", "','") & "')"
If clpCode(3) <> "" Then QStr = QStr & " AND ED_REGION='" & clpCode(3) & "'"
If clpCode(4) <> "" Then QStr = QStr & " AND ED_ADMINBY='" & clpCode(4) & "'"
If clpCode(5) <> "" Then QStr = QStr & " AND ED_SECTION='" & clpCode(5) & "'"
If clpCode(6) <> "" Then QStr = QStr & " AND ED_PT in ('" & Replace(clpCode(6), ",", "','") & "')"
If elpEEID.Text <> "" Then QStr = QStr & " AND ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
If WithAtt Then
    If IsDate(dlpDateRange(0)) Then QStr = QStr & " AND AD_DOA>=" & Date_SQL(DATE1)
    If IsDate(dlpDateRange(1)) Then QStr = QStr & " AND AD_DOA<=" & Date_SQL(DATE1)
'    If clpAtt <> "" Then QStr = QStr & " AND ES_CTYPE IN ('" & Replace(clpAtt, ",", "','") & "') "
'    If clpPayP <> "" Then QStr = QStr & " AND ES_CRSCODE IN ('" & Replace(clpPayP, ",", "','") & "') "
'    If txtShift <> "" Or clpJob <> "" Or clpPosGroup <> "" Then
'        QStr = QStr & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
'        If txtShift <> "" Then QStr = QStr & " AND JH_SHIFT='" & txtShift & "'"
'        If clpJob <> "" Then QStr = QStr & " AND JH_JOB='" & clpJob & "'"
'        If clpPosGroup <> "" Then QStr = QStr & " AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD='" & clpPosGroup & "')"
'        QStr = QStr & ")"
'    End If
End If

getWSQLQ = QStr
End Function

Private Sub Export_Employment_Equity()
    Dim rsEmpEquity As New ADODB.Recordset
    Dim rsCompInfo As New ADODB.Recordset
    Dim exApp As Excel.Application
    Dim exBook As Excel.Workbook
    Dim exSheet As Excel.Worksheet
    Dim SQLQ, sSQLQ As String
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xStatus As String
    Dim xRow, xCol As Long
    Dim I, totNum, xEmpDisp
    Dim noStatus As Boolean
    Dim xExcelRptPath  As String

    On Error GoTo Export_Employment_Equity_Err
    
    'Ticket #22034 - Get Excel reports path
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If

    'Get Employees to display
    SQLQ = "SELECT * FROM HREMPEQU "
    SQLQ = SQLQ & " WHERE 1 = 1 "
    
    'Category
    If Len(clpCode(6).Text) > 0 Then
        SQLQ = SQLQ & " AND EQ_EEPT in ('" & Replace(clpCode(6), ",", "','") & "')"
    End If
    
    'Region
    If Len(clpCode(1).Text) > 0 Then
        SQLQ = SQLQ & " AND EQ_REGION in ('" & Replace(clpCode(1), ",", "','") & "')"
    End If
    
    'Type
    If cmbType.ListIndex <> 2 Then
        SQLQ = SQLQ & " AND EQ_TYPE='" & UCase(Left(cmbType.Text, 1)) & "'"
    End If
    
    'Employee
    If Len(elpEEID.Text) > 0 Then
        SQLQ = SQLQ & " AND EQ_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & ") "
    End If
    
    'Plan
    If Len(clpPlanNbr.Text) > 0 Then
        SQLQ = SQLQ & " AND EQ_PLAN='" & clpPlanNbr & "'"
    End If
    
    'Province
    If Len(clpProv.Text) > 0 Then
        SQLQ = SQLQ & " AND EQ_PROV='" & txtProv.Text & "'"
    End If
    
    If Len(sSQLQ) > 0 Or IsDate(dlpDateRange(0)) Or IsDate(dlpDateRange(1)) Then
        sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "HREMP.", "")
        SQLQ = SQLQ & " AND EQ_EMPNBR IN (SELECT ED_EMPNBR FROM qry_Current_Active_Term_Employees WHERE 1 = 1 "
        
        If Len(sSQLQ) > 0 Then SQLQ = SQLQ & " AND " & sSQLQ
        
        If IsDate(dlpDateRange(0)) Then SQLQ = SQLQ & " AND ((ACTTERM = 'A' AND ED_DOH>=" & Date_SQL(dlpDateRange(0))
        If IsDate(dlpDateRange(1)) Then SQLQ = SQLQ & " AND ED_DOH<=" & Date_SQL(dlpDateRange(1)) & ")"
        If IsDate(dlpDateRange(0)) Then SQLQ = SQLQ & " OR (ACTTERM = 'T' AND TERM_DOT>=" & Date_SQL(dlpDateRange(0))
        If IsDate(dlpDateRange(1)) Then SQLQ = SQLQ & " AND TERM_DOT<=" & Date_SQL(dlpDateRange(1)) & "))"
        
        SQLQ = SQLQ & ")"
    End If
    
    SQLQ = SQLQ & " ORDER BY EQ_EMPNBR"
    rsEmpEquity.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    
    If Not rsEmpEquity.EOF Then
        totNum = rsEmpEquity.RecordCount: I = 0
        xEmpDisp = 0
        
        rsEmpEquity.MoveFirst

        xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpEquityTmp.xls"
        
        'Ticket #22034 - May need to save the report in different path
        'xlsFileMat = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "EmpEquityRpt" & Trim(glbUserID) & ".xls"
        xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "EmpEquityRpt" & Trim(glbUserID) & ".xls"
    
        If Dir(xlsFileTmp) = "" Then
            MsgBox "There is no " & xlsFileTmp
            Exit Sub
        End If
        If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
    
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
    
        FileCopy xlsFileTmp, xlsFileMat
    
        Screen.MousePointer = HOURGLASS
    
        'Create new WorkBook of Excel
        Set exApp = CreateObject("Excel.Application")
        Set exBook = exApp.Workbooks.Open(xlsFileMat)
        Set exSheet = exBook.Worksheets(1)
    
        SQLQ = "SELECT PC_NAME FROM HRPARCO"
        rsCompInfo.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        Dim companyname As String
        If Not rsCompInfo.EOF Then
            rsCompInfo.MoveFirst
            companyname = rsCompInfo("PC_NAME")
        End If
    
        rsCompInfo.Close
        
        exSheet.Cells(1, 3) = companyname
        
        exSheet.Cells(1, 1) = "Date: " & Format(Now, "mm/dd/yyyy")
        exSheet.Cells(2, 1) = "Time: " & Time$
                
        xRow = 5
        'Columns: 1 - Emp #, 2 - CMA, 3 - Prov, 4 - NOC, etc.
        Do While Not rsEmpEquity.EOF
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            noStatus = True
                        
            'Display field values
            exSheet.Cells(xRow, 1) = rsEmpEquity("EQ_EMPNBR")
            exSheet.Cells(xRow, 2) = rsEmpEquity("EQ_ORGT1")
            exSheet.Cells(xRow, 3) = rsEmpEquity("EQ_PROV")
            exSheet.Cells(xRow, 4) = rsEmpEquity("EQ_NOGC")
            exSheet.Cells(xRow, 5) = rsEmpEquity("EQ_NAICS")
            exSheet.Cells(xRow, 6) = rsEmpEquity("EQ_EEPT")
            exSheet.Cells(xRow, 8) = rsEmpEquity("EQ_EESEX")
            If rsEmpEquity("EQ_TYPE") = "T" Then
                exSheet.Cells(xRow, 9) = GetTermSHData(rsEmpEquity("EQ_EMPNBR"), "SH_SALARY", "")
            ElseIf IsDate(rsEmpEquity("EQ_DOT")) Then
                exSheet.Cells(xRow, 9) = GetTermSHData(rsEmpEquity("EQ_EMPNBR"), "SH_SALARY", "")
            Else
                exSheet.Cells(xRow, 9) = GetSHData(rsEmpEquity("EQ_EMPNBR"), "SH_SALARY", "")
            End If
            exSheet.Cells(xRow, 10) = rsEmpEquity("EQ_ABORYN")
            exSheet.Cells(xRow, 11) = rsEmpEquity("EQ_VMYN")
            exSheet.Cells(xRow, 12) = rsEmpEquity("EQ_DISAYN")
            If IsDate(rsEmpEquity("EQ_DOT")) Then
                exSheet.Cells(xRow, 13) = Format(GetEmpData(rsEmpEquity("EQ_EMPNBR"), "ED_DOH", "T", ""), "mm/dd/yyyy")
            Else
                exSheet.Cells(xRow, 13) = Format(GetEmpData(rsEmpEquity("EQ_EMPNBR"), "ED_DOH", "A", ""), "mm/dd/yyyy")
            End If
            exSheet.Cells(xRow, 14) = Format(rsEmpEquity("EQ_DOT"), "mm/dd/yyyy")
            exSheet.Cells(xRow, 16) = rsEmpEquity("EQ_REGION")
            
            exSheet.Range("A" & xRow & ":P" & xRow).Borders(xlInsideVertical).LineStyle = xlThin
            exSheet.Range("A" & xRow & ":P" & xRow).Borders(xlEdgeBottom).LineStyle = xlDot
            
            xEmpDisp = xEmpDisp + 1
            
            xRow = xRow + 1
Next_Employee:
            rsEmpEquity.MoveNext
        Loop
        
        If xRow > 5 Then
            exSheet.Range("A5:A" & xRow - 1).Borders(xlEdgeLeft).Weight = xlThin
            exSheet.Range("P5:P" & xRow - 1).Borders(xlEdgeRight).Weight = xlThin
            exSheet.Range("A" & xRow - 1 & ":P" & xRow - 1).Borders(xlEdgeBottom).Weight = xlThin
            exSheet.Range("A" & xRow - 1 & ":P" & xRow - 1).Borders(xlEdgeBottom).LineStyle = xlThin
        End If
        
        exSheet.Cells(xRow + 2, 2) = "Total Number of Employees : " & xEmpDisp
        exSheet.Rows(xRow + 2).Font.Bold = True

        exBook.Save
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing

    
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    Else
        MsgBox "No data for this selection", vbOKOnly, "Employment Equity Report"
    End If
    rsEmpEquity.Close
            
Exit Sub

Export_Employment_Equity_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "Employment Equity Report", "SELECT")
'Resume Next
Set exSheet = Nothing
Set exBook = Nothing
Set exApp = Nothing
            
End Sub

Function LanchXlsW98(xFileName)
On Error GoTo Error_Deal
    LanchXlsW98 = False
    Shell "Start " & GetShortName(xFileName)
    LanchXlsW98 = True
Exit Function
Error_Deal:

End Function

Private Function GetEmpData(EmpNbr, Field As String, xActTermAT, DEFAULT)
    Dim rsEE As New ADODB.Recordset
    
    If xActTermAT = "A" Then
        rsEE.Open "SELECT " & Field & " FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
        GetEmpData = DEFAULT
    ElseIf xActTermAT = "T" Then
        rsEE.Open "SELECT " & Field & " FROM Term_HREMP WHERE ED_EMPNBR=" & EmpNbr & " ORDER BY TERM_SEQ DESC", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

        GetEmpData = DEFAULT
    End If
    If Not rsEE.EOF Then
        If Not IsNull(rsEE(Field)) Then GetEmpData = rsEE(Field)
    End If

End Function
