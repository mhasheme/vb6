VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRSalVacPrcIncr 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Salary/Vacation % Increase Report"
   ClientHeight    =   7470
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   9690
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
   ScaleHeight     =   7470
   ScaleWidth      =   9690
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1875
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3700
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   390
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
      Left            =   1560
      TabIndex        =   9
      Tag             =   "00-Enter Section Code"
      Top             =   3360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   1560
      TabIndex        =   8
      Tag             =   "00-Enter Administered By Code"
      Top             =   3030
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
      Left            =   1560
      TabIndex        =   7
      Tag             =   "00-Enter Region Code"
      Top             =   2700
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1560
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   2040
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
      Index           =   2
      Left            =   1560
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1710
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
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1380
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
      Top             =   1050
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
      Top             =   720
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
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "Second level of grouping records"
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
      Left            =   1890
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Tag             =   "First Level of grouping records"
      Top             =   5325
      Width           =   2325
   End
   Begin Threed.SSCheck chkShowException 
      Height          =   225
      Left            =   1860
      TabIndex        =   13
      Tag             =   "If X-Show Exception only"
      Top             =   4500
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   " Exception from rules only"
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
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2370
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
      Top             =   5640
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
      Left            =   3210
      TabIndex        =   12
      Tag             =   "40-Date upto and including this date forward"
      Top             =   4050
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
      Left            =   1560
      TabIndex        =   11
      Tag             =   "40-Date from and including this date forward"
      Top             =   4050
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
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
      Left            =   120
      TabIndex        =   31
      Top             =   4095
      Visible         =   0   'False
      Width           =   1095
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
      TabIndex        =   30
      Top             =   3745
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
      TabIndex        =   29
      Top             =   2085
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
      TabIndex        =   28
      Top             =   3405
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
      TabIndex        =   27
      Top             =   2745
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
      TabIndex        =   26
      Top             =   3075
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
      TabIndex        =   25
      Top             =   1095
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
      TabIndex        =   24
      Top             =   5070
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
      TabIndex        =   23
      Top             =   5700
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
      TabIndex        =   22
      Top             =   5385
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
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   2415
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
      TabIndex        =   19
      Top             =   1755
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
      TabIndex        =   18
      Top             =   1425
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
      TabIndex        =   17
      Top             =   765
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
      TabIndex        =   16
      Top             =   435
      Width           =   555
   End
End
Attribute VB_Name = "frmRSalVacPrcIncr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HisSQL As String

Private Sub chkShowException_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm("Salary/Vacation % Increase Report Criteria", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    X% = Cri_SetAll()
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
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
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

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
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
      Call set_PrintState(True)
End If

Exit Sub

CRW_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "ENTITLEMENTS", "VIEW")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub comEmpType_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    
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

Private Sub Cri_Shift()
Dim EECri As String, OneSet%, X%

If Len(txtShift.Text) < 1 Then Exit Sub
EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"
HisSQL = HisSQL & " AND (HREMP.ED_SHIFT ='" & txtShift.Text & "')"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
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
    
    If glbLinamar And (strCd$ = "HREMP.ED_REGION" Or strCd$ = "HREMP.ED_SECTION") Then
        CodeCri = "(({" & strCd$ & "} = '" & clpDiv.Text & clpCode(intIdx%).Text & "') or ({" & strCd$ & "} = 'ALL" & clpCode(intIdx%).Text & "') )"
        HisSQL = HisSQL & " AND ((" & strCd$ & "'" & clpDiv.Text & clpCode(intIdx%).Text & "') or (" & strCd$ & " = 'ALL" & clpCode(intIdx%).Text & "') )"
    Else
        CodeCri = "({" & strCd$ & "} in  ['" & Replace(clpCode(intIdx%).Text, ",", "','") & "'])"
        HisSQL = HisSQL & " AND (" & strCd$ & " IN ('" & Replace(clpCode(intIdx%).Text, ",", "','") & "'))"
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
    DivCri = "({HREMP.ED_DIV} in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
    HisSQL = HisSQL & " AND (HREMP.ED_DIV in ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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
    HisSQL = HisSQL & " AND (HREMP.ED_EMPNBR IN (" & getEmpnbr(elpEEID.Text) & "))"
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
Dim EECri As String, OneSet%, X%

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "{HREMP.ED_PT} in ['" & Replace(clpPT.Text, ",", "','") & "']"
HisSQL = HisSQL & " AND (HREMP.ED_PT in ['" & Replace(clpPT.Text, ",", "','") & "'])"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If

glbiOneWhere = True

End Sub

Private Function Cri_SetAll()
Dim X%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
HisSQL = ""
' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_PT
Call Cri_Shift
Call Cri_EE

'Call function to populate the working table with records for the report.
Call Populate_SalVac_BasedOn_Rule

' report name
  If comGroup(0) <> "(none)" Then
    strRName$ = glbIHRREPORTS & "rzSalVacPrc.rpt"
  Else
    strRName$ = glbIHRREPORTS & "rzSalVacPrc1.rpt"
  End If
  Me.vbxCrystal.ReportFileName = strRName$
  
  ' set to sorting/grouping criteria
  X% = Cri_Sorts()   ' returns number of sections formated
  
'  If chkShowException.Value = False Then
'    If Len(glbstrSelCri) >= 0 Then
'       glbstrSelCri = "(" & glbstrSelCri & " And " & "len({HREMP.ED_EMAIL}) > 0 )"
'    Else
'       glbstrSelCri = "(len({HREMP.ED_EMAIL}) > 0 )"
'    End If
'  End If
  
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Salary/Vacation % Incr", "HR_SALARY_INCR", "Select")
Cri_SetAll = False
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
'first set primary grouping
Y% = 0
grpField$ = getEGroup(comGroup(0).Text)

Call setRptLabel(Me, 0)
If comGroup(0) = "(none)" Then Exit Function
    
Y% = X% + 1
dscGroup$ = comGroup(X%).Text
dscGroup$ = "descGroup" & CStr(Y%) & "= '" & dscGroup$ & "'"
Me.vbxCrystal.Formulas(X%) = dscGroup$

grpCond$ = "GROUP" & CStr(Y%) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(X%) = grpCond$

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

Cri_Sorts = z% ' next section number to format

End Function

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

If Not clpPT.ListChecker Then
'If Len(clpPT) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

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

glbOnTop = "FRMRSALVACPRCINCR"

Screen.MousePointer = HOURGLASS
frmRSalVacPrcIncr.Caption = "Salary/Vacation % Increase Report"

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

Call comGrpLoad
Call setRptCaption(Me)

If glbLinamar Then clpCode(3).MaxLength = 8
If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
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

Private Sub Populate_SalVac_BasedOn_Rule()
    Dim rsSalVacWrk As New ADODB.Recordset
    Dim rsSalRule As New ADODB.Recordset
    Dim rsVacRule As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsEMP As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim SQLQ As String
    Dim xSalary, xActSal, xActSalGrd, xSalGrd
    Dim xVacPctg, xActVacPctg
    Dim xDiv, xDept, xORG, xLoc, xEMP, xPT, xGRPCD, xSec, xJob
    Dim xSalType As String
    Dim z As Integer
    
    'Delete from the HRSALVACWRK table for this user
    gdbAdoIhr001.Execute "DELETE FROM HRSALVACWRK " & in_SQL(glbIHRDBW) & " WHERE WRKEMP='" & glbUserID & "'"
    
    'Open working table recordset
    SQLQ = "SELECT * FROM HRSALVACWRK WHERE 1=2"
    rsSalVacWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    'Get the Total Seniority Hours from HR_ATTENDANCE for all the employees
    'Ticket #15276 - Include HR_ATTENDANCE_HISTORY table as well in the calculation for total seniority hours
    SQLQ = "SELECT EMPNBR, SUM(TOT_SEN_HRS) AS TOT_SEN_HRS FROM "
    SQLQ = SQLQ & " (SELECT AD_EMPNBR AS EMPNBR, SUM(AD_HRS) AS TOT_SEN_HRS FROM HR_ATTENDANCE WHERE"
    SQLQ = SQLQ & " AD_SEN<>0 "
    SQLQ = SQLQ & " AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & HisSQL & ")"
    SQLQ = SQLQ & " GROUP BY AD_EMPNBR"
    SQLQ = SQLQ & " UNION "
    SQLQ = SQLQ & " SELECT AH_EMPNBR AS EMPNBR, SUM(AH_HRS) AS TOT_SEN_HRS FROM HR_ATTENDANCE_HISTORY WHERE"
    SQLQ = SQLQ & " AH_SEN<>0 "
    SQLQ = SQLQ & " AND AH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & HisSQL & ")"
    SQLQ = SQLQ & " GROUP BY AH_EMPNBR) AS HR_ATTENDANCE"
    SQLQ = SQLQ & " GROUP BY EMPNBR"
    
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsAttend.EOF Then
        rsAttend.MoveFirst
        
        Do While Not rsAttend.EOF
            'Initialize
            xActSal = 0
            xSalary = 0
            xActVacPctg = 0
            xVacPctg = 0
            xSalType = ""
            xActSalGrd = 0
            xSalGrd = 0
            xJob = ""
                    
            'Get Actual Vacation %
            SQLQ = "SELECT ED_EMPNBR, ED_VACPC FROM HREMP WHERE ED_EMPNBR =" & rsAttend("EMPNBR")
            rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEMP.EOF Then
                xActVacPctg = rsEMP("ED_VACPC") & ""
            End If
            rsEMP.Close
            Set rsEMP = Nothing
            
            'Get rule based Salary
            SQLQ = "SELECT * FROM HR_SALARY_INCR"
            rsSalRule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSalRule.EOF Then
                rsSalRule.MoveFirst
                
                Do While Not rsSalRule.EOF
                    xDiv = rsSalRule("SL_DIV") & ""
                    xDept = rsSalRule("SL_DEPT") & ""
                    xORG = rsSalRule("SL_ORG") & ""
                    xLoc = rsSalRule("SL_LOC") & ""
                    xEMP = rsSalRule("SL_EMP") & ""
                    xPT = rsSalRule("SL_PT") & ""
                    xGRPCD = rsSalRule("SL_GRPCD") & ""
                    xSec = rsSalRule("SL_SECTION") & ""
                
                    'Select employees who belong to this rule
                    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & rsAttend("EMPNBR")
                    If Len(xDiv) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & xDiv & "'"
                    If Len(xDept) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO = '" & xDept & "'"
                    If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "'"
                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_LOC = '" & xLoc & "'"
                    If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "'"
                    If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "'"
                    If Len(xGRPCD) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD = '" & xGRPCD & "'))"
                    If glbLinamar Then  'Hemu - Ticket #12494 - Section is actually Salary Distribution field for them on Entitl. Mstr
                        If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SALDIST = '" & xSec & "' "
                    Else
                        If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SECTION = '" & xSec & "'"
                    End If
                    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEMP.EOF Then
                        'Compare Total Seniority Hours with the range of the rule
                        If rsAttend("TOT_SEN_HRS") >= rsSalRule("SL_FROM_HRS") And rsAttend("TOT_SEN_HRS") <= rsSalRule("SL_TO_HRS") Then
                            'Get the Salary/Grid Step based on the rule this employee should belong to.
                            xSalType = rsSalRule("SL_TYPE")
                            If xSalType = "A" Then
                                xSalary = rsSalRule("SL_SALARY")
                            Else
                                xSalGrd = rsSalRule("SL_SALARY")
                            End If
                        End If
                    End If
                    rsEMP.Close
                    Set rsEMP = Nothing
                    
                    rsSalRule.MoveNext
                Loop
            End If
            rsSalRule.Close
            
            'Get Actual Current Salary/Grade of the employee
            SQLQ = "SELECT SH_EMPNBR,SH_SALARY,SH_GRADE,SH_JOB FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & rsAttend("EMPNBR") & " AND SH_CURRENT <>0"
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                'If xSalType = "A" Then
                    xActSal = rsSal("SH_SALARY")
                    xActSalGrd = rsSal("SH_GRADE")
                'Else
                '    xActSal = rsSal("SH_SALARY")    'rsSal("SH_GRADE")
                '    xActSalGrd = rsSal("SH_GRADE")  'rsSal("SH_SALARY")
                'End If
                xJob = rsSal("SH_JOB")  'to get the salary/grade from HRJob
            End If
            rsSal.Close
            Set rsSal = Nothing
    
            'Get the Salary/Grade of the Rule from the HRJOB for this employee's Position
            If xSalType = "A" Then
                rsJOB.Open "SELECT * FROM HRJOB WHERE JB_CODE = '" & xJob & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsJOB.EOF Then
                    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
                    'For z = 1 To 11
                    'For z = 1 To 15
                    For z = 1 To 20
                        If rsJOB("JB_S" & z) = xSalary Then
                            xSalGrd = z
                        Else
                            'Ticket #15276 - if Grid has 0 or blank then get last valid grid step which has > 0 salary
                            If rsJOB("JB_S" & z) <> 0 And rsJOB("JB_S" & z) <> "" Then
                                xSalGrd = z
                            End If
                        End If
                    Next z
                End If
                rsJOB.Close
                Set rsJOB = Nothing
            Else
                If xSalType <> "" And xSalGrd <> 0 Then
                    rsJOB.Open "SELECT * FROM HRJOB WHERE JB_CODE = '" & xJob & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsJOB.EOF Then
                        xSalary = rsJOB("JB_S" & CInt(xSalGrd))
                        'Ticket #15276 - if Salary retrieve is 0 or blank then get last valid grid step salary
                        If xSalary = 0 Or xSalary = "" Then
                            For z = CInt(xSalGrd) To 1 Step -1
                                If rsJOB("JB_S" & CInt(z)) > 0 Then
                                    xSalary = rsJOB("JB_S" & CInt(z))
                                    Exit For
                                End If
                            Next z
                        End If
                    End If
                    rsJOB.Close
                    Set rsJOB = Nothing
                End If
            End If
            
            'Get rule based Vacation %
            SQLQ = "SELECT * FROM HR_VACATION_INCR"
            rsVacRule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsVacRule.EOF Then
                rsVacRule.MoveFirst
                
                Do While Not rsVacRule.EOF
                    xDiv = rsVacRule("VC_DIV") & ""
                    xDept = rsVacRule("VC_DEPT") & ""
                    xORG = rsVacRule("VC_ORG") & ""
                    xLoc = rsVacRule("VC_LOC") & ""
                    xEMP = rsVacRule("VC_EMP") & ""
                    xPT = rsVacRule("VC_PT") & ""
                    xGRPCD = rsVacRule("VC_GRPCD") & ""
                    xSec = rsVacRule("VC_SECTION") & ""
                
                    'Select employees who belong to this rule
                    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & rsAttend("EMPNBR")
                    If Len(xDiv) > 0 Then SQLQ = SQLQ & " AND ED_DIV = '" & xDiv & "'"
                    If Len(xDept) > 0 Then SQLQ = SQLQ & " AND ED_DEPTNO = '" & xDept & "'"
                    If Len(xORG) > 0 Then SQLQ = SQLQ & " AND ED_ORG = '" & xORG & "'"
                    If Len(xLoc) > 0 Then SQLQ = SQLQ & " AND ED_LOC = '" & xLoc & "'"
                    If Len(xEMP) > 0 Then SQLQ = SQLQ & " AND ED_EMP = '" & xEMP & "'"
                    If Len(xPT) > 0 Then SQLQ = SQLQ & " AND ED_PT = '" & xPT & "'"
                    If Len(xGRPCD) > 0 Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB IN (SELECT JB_CODE FROM HRJOB WHERE JB_GRPCD = '" & xGRPCD & "'))"
                    If glbLinamar Then  'Hemu - Ticket #12494 - Section is actually Salary Distribution field for them on Entitl. Mstr
                        If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SALDIST = '" & xSec & "' "
                    Else
                        If Len(xSec) > 0 Then SQLQ = SQLQ & " AND ED_SECTION = '" & xSec & "'"
                    End If
                    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsEMP.EOF Then
                        'Compare Total Seniority Hours with the range of the rule
                        If rsAttend("TOT_SEN_HRS") >= rsVacRule("VC_FROM_HRS") And rsAttend("TOT_SEN_HRS") <= rsVacRule("VC_TO_HRS") Then
                            'Get the Vacation % based on the rule this employee should belong to.
                            xVacPctg = rsVacRule("VC_PCT")
                        End If
                    End If
                    rsEMP.Close
                    Set rsEMP = Nothing
                    
                    rsVacRule.MoveNext
                Loop
            End If
            rsVacRule.Close
            
            
            If chkShowException.Value = False Or (chkShowException.Value = True And (xActSal <> xSalary Or xActVacPctg <> xVacPctg)) Then
                'Insert records in the working table for the report
                rsSalVacWrk.AddNew
                rsSalVacWrk("SV_EMPNBR") = rsAttend("EMPNBR")
                rsSalVacWrk("SV_SEN_HRS") = rsAttend("TOT_SEN_HRS")
                If Not IsNull(xActSal) And xActSal <> "" Then
                    rsSalVacWrk("SV_ACT_SAL") = xActSal
                End If
                If Not IsNull(xActSalGrd) And xActSalGrd <> "" Then
                    rsSalVacWrk("SV_ACT_SAL_GRD") = xActSalGrd
                End If
                
                If Not IsNull(xActVacPctg) And xActVacPctg <> "" Then
                    rsSalVacWrk("SV_ACT_VACPCT") = xActVacPctg
                End If
                
                If Not IsNull(xSalary) And xSalary <> "" Then
                    rsSalVacWrk("SV_SAL") = xSalary
                End If
                If Not IsNull(xSalGrd) And xSalGrd <> "" Then
                    rsSalVacWrk("SV_SAL_GRD") = xSalGrd
                End If
                
                If Not IsNull(xVacPctg) And xVacPctg <> "" Then
                    rsSalVacWrk("SV_VACPCT") = xVacPctg
                End If
                rsSalVacWrk("SV_SAL_TYPE") = xSalType
                rsSalVacWrk("WRKEMP") = glbUserID
                rsSalVacWrk("SV_LUSER") = glbUserID
                rsSalVacWrk("SV_LDATE") = Date
                rsSalVacWrk("SV_LTIME") = Time$
                                
                rsSalVacWrk.Update
            End If
            
            rsAttend.MoveNext
        Loop
    End If
    rsAttend.Close
    rsSalVacWrk.Close
    
End Sub
