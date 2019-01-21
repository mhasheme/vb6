VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRFCOCDiscipline 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "COC Discipline Form"
   ClientHeight    =   7110
   ClientLeft      =   375
   ClientTop       =   915
   ClientWidth     =   9795
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
   ScaleHeight     =   7110
   ScaleWidth      =   9795
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1875
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "00-Employee Position Shift"
      Top             =   3795
      Visible         =   0   'False
      Width           =   450
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   510
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
      Top             =   3480
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
      Top             =   3150
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
      Top             =   2820
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
      Top             =   2160
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
      Top             =   1830
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1170
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
      Top             =   840
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
      TabIndex        =   16
      Tag             =   "Second level of grouping records"
      Top             =   6360
      Visible         =   0   'False
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
      TabIndex        =   15
      Tag             =   "First Level of grouping records"
      Top             =   6045
      Visible         =   0   'False
      Width           =   2325
   End
   Begin Threed.SSCheck chkShowEmp 
      Height          =   225
      Left            =   4560
      TabIndex        =   13
      Tag             =   "If X-Show All Employees"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2835
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1323
      _StockProps     =   78
      Caption         =   " Show All Employees"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27
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
      Top             =   2490
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
      Top             =   6600
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   6
      Left            =   1560
      TabIndex        =   12
      Tag             =   "01-Counselling Reason- Code"
      Top             =   4460
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CERE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   7
      Left            =   1560
      TabIndex        =   11
      Tag             =   "01-Counselling Type- Code"
      Top             =   4120
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CETY"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   32
      Top             =   4475
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   31
      Top             =   4135
      Width           =   735
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
      Top             =   3840
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
      Top             =   2205
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
      Top             =   3525
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
      Top             =   2865
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
      Top             =   3195
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
      Top             =   1215
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
      Top             =   5835
      Visible         =   0   'False
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
      Top             =   6420
      Visible         =   0   'False
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
      Top             =   6105
      Visible         =   0   'False
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
      Top             =   2535
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
      Top             =   1875
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
      Top             =   1545
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
      Top             =   885
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
      TabIndex        =   14
      Top             =   555
      Width           =   555
   End
End
Attribute VB_Name = "frmRFCOCDiscipline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Excel2007 = 12
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Private Sub chkShowEmp_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then

    If Not PrtForm("Attendance Sign In Form Criteria", Me) Then Exit Sub
    
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
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    'Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
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
    comGroup(1).AddItem "Employee Number"
    comGroup(1).ListIndex = 0

End Sub

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

Private Sub Cri_PT()
Dim EECri As String, OneSet%, X%

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
Dim X%, strRName$

Cri_SetAll = False
On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_PT
Call Cri_Shift
Call Cri_EE

Call Showa_COC_Discipline_Form

' report name
'If comGroup(0) <> "(none)" Then
'    strRName$ = glbIHRREPORTS & "rzemail.rpt"
'Else
'    strRName$ = glbIHRREPORTS & "rzemail1.rpt"
'End If
'Me.vbxCrystal.ReportFileName = strRName$
'set to sorting/grouping criteria
'X% = Cri_Sorts()   ' returns number of sections formated
'If Len(glbstrSelCri) >= 0 Then
'    Me.vbxCrystal.SelectionFormula = glbstrSelCri
'End If
'Me.vbxCrystal.Connect = RptODBC_SQL

  
Cri_SetAll = True
Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cri_SetAll", "COC Discipline Form", "Select")
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

If comGroup(0) = "(none)" Then
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$

    Exit Function
End If
    
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
    Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
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

If Trim(clpCode(7).Text) = "" Then
    MsgBox "Type is a required field."
    clpCode(7).SetFocus
    Exit Function
End If

If Not clpCode(7).ListChecker Then Exit Function

If Trim(clpCode(6).Text) = "" Then
    MsgBox "Reason is a required field."
    clpCode(6).SetFocus
    Exit Function
End If

If Not clpCode(6).ListChecker Then Exit Function

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMRFCOCDISCIPLINE"

Screen.MousePointer = HOURGLASS

frmRFCOCDiscipline.Caption = "COC Discipline Form"

If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If

'Call comGrpLoad
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

Private Sub Showa_COC_Discipline_Form()
    Dim rsCounsel As New ADODB.Recordset
    Dim rsCounselHist As New ADODB.Recordset
    Dim SQLQ As String
    Dim sSQLQ As String

    Dim exApp As Object
    Dim exBook As Object
    Dim exSheet As Object
    Dim xlsFileTmp As String
    Dim xlsFileMat As String
    Dim xRow As Long
    Dim I, totNum
    Dim lstEmpNo
    Dim appVerInt As Double
    Dim Total As Integer
    Dim xCounselHist As String
    Dim xCounselHistCount As Integer

    On Error GoTo Showa_COC_Discipline_Form_Err
    
    Screen.MousePointer = HOURGLASS
    
    'Ticket #22034 - Get Excel reports path
    Dim xExcelRptPath  As String
    If gsTRAININGMATRIX Then
        xExcelRptPath = GetComPreferEmail("TRAININGMATRIX")
    End If
    If Len(xExcelRptPath) = 0 Then
        xExcelRptPath = glbIHRREPORTS
    End If
        
    sSQLQ = Replace(Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")"), "Uppercase", "upper")
    
    SQLQ = "SELECT CL_EMPNBR FROM HR_COUNSEL"
    SQLQ = SQLQ & " WHERE CL_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 AND " & sSQLQ & ")"
    SQLQ = SQLQ & " AND CL_TYPE = '" & clpCode(7).Text & "'"
    SQLQ = SQLQ & " AND CL_REASON = '" & clpCode(6).Text & "'"
    SQLQ = SQLQ & " GROUP BY CL_EMPNBR"
    
    'Call WriteFile("SQL1=" & SQLQ)
        
    Total = 0
    lstEmpNo = 0
    
    rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If Not rsCounsel.EOF Then
        rsCounsel.MoveFirst
        totNum = rsCounsel.RecordCount: I = 0
                
        xRow = 4
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(0).FloodPercent = 0
                
        Do While Not rsCounsel.EOF
            
            If (I / totNum) <= 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / totNum) * 100
                I = I + 1
            End If
            DoEvents
            
            'File to export to
            xlsFileTmp = glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "SN2454COCDiscTmp.xls"
            xlsFileMat = xExcelRptPath & IIf(Right(xExcelRptPath, 1) = "\", "", "\") & "COCDiscForm_" & Trim(glbUserID) & "_" & rsCounsel("CL_EMPNBR") & ".xls"
        
            If Dir(xlsFileTmp) = "" Then
                Screen.MousePointer = DEFAULT
                MsgBox "There is no " & xlsFileTmp
                Exit Sub
            End If
            If (Dir(xlsFileMat)) <> "" Then Kill xlsFileMat
                    
            FileCopy xlsFileTmp, xlsFileMat
            
            'Create new WorkBook of Excel
            Set exApp = CreateObject("Excel.Application")
            Set exBook = exApp.Workbooks.Open(xlsFileMat)
            Set exSheet = exBook.Worksheets(1)
    
            'Ticket #22166
            appVerInt = Split(exApp.Version, ".")(0)
            If appVerInt - Excel2007 >= 0 Then
                'exApp.ActiveWorkbook.SaveAs (sXLS), 56
                exApp.DisplayAlerts = False
                exBook.SaveAs (xlsFileMat), 56
                exApp.DisplayAlerts = True
            Else
                'exApp.ActiveWorkbook.SaveAs (sXLS), 43
                exApp.DisplayAlerts = False
                exBook.SaveAs (xlsFileMat), 43
                exApp.DisplayAlerts = True
            End If
            
            xCounselHist = ""
            xCounselHistCount = 0
            'Retrieve the Counsel History of the employee
            SQLQ = "SELECT * FROM HR_COUNSEL"
            SQLQ = SQLQ & " WHERE CL_EMPNBR = " & rsCounsel("CL_EMPNBR")
            SQLQ = SQLQ & " AND CL_TYPE = '" & clpCode(7).Text & "'"
            SQLQ = SQLQ & " AND CL_REASON = '" & clpCode(6).Text & "'"
            SQLQ = SQLQ & " ORDER BY CL_EMPNBR, CL_TYPE, CL_REASON, CL_LEVEL DESC"
            rsCounselHist.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsCounselHist.EOF Then
                '# of Counsel History
                xCounselHistCount = rsCounselHist.RecordCount
                rsCounselHist.MoveFirst
                
                Do While Not rsCounselHist.EOF
                    'Print field values
                    exSheet.Cells(1, 3) = Format(Now, "mm/dd/yyyy")
                    exSheet.Cells(1, 5) = rsCounsel("CL_EMPNBR")
                    exSheet.Cells(1, 7) = Format(GetEmpData(rsCounsel("CL_EMPNBR"), "ED_DOH"), "mm/dd/yyyy")
                    exSheet.Cells(2, 3) = GetEmpData(rsCounsel("CL_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsCounsel("CL_EMPNBR"), "ED_SURNAME")
                    exSheet.Cells(2, 7) = Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy")
                                                                
                    If rsCounselHist("CL_LEVEL") = 1 Then
                        exSheet.Range("B4:B4").Interior.Color = vbRed   'Ticket #25826 - Client want Red instead of Yellow
                    ElseIf rsCounselHist("CL_LEVEL") = 2 Then
                        exSheet.Range("C4:C4").Interior.Color = vbRed
                    ElseIf rsCounselHist("CL_LEVEL") = 3 Then
                        exSheet.Range("D4:D4").Interior.Color = vbRed
                    ElseIf rsCounselHist("CL_LEVEL") = 4 Then
                        exSheet.Range("E4:E4").Interior.Color = vbRed
                    ElseIf rsCounselHist("CL_LEVEL") = 5 Then
                        exSheet.Range("F4:F4").Interior.Color = vbRed
                    ElseIf rsCounselHist("CL_LEVEL") = 6 Then
                        exSheet.Range("G4:G4").Interior.Color = vbRed
                    End If
                    
                    'Stream
                    'Ticket #27619 - Concatenate Comments to the Stream.
                    'exSheet.Cells(11, 2) = Get_StreamDesc(rsCounselHist("CL_STREAM"))
                    exSheet.Cells(11, 2).Value = Get_StreamDesc(rsCounselHist("CL_STREAM")) & vbCrLf & rsCounselHist("CL_COMMENTS")
                    
                    'Previously higher level Incident Date and Level # - C3 and F3
                    If xCounselHistCount = 1 Then
                        exSheet.Cells(3, 3) = "NA" 'Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy")   'Ticket #25826 - Client says to show NA if one and only record
                        exSheet.Cells(3, 6) = rsCounselHist("CL_LEVEL")
                    Else
                        'Check the Incident Date of next highest level
                        rsCounselHist.MoveNext
                        exSheet.Cells(3, 3) = Format(rsCounselHist("CL_INCDATE"), "dd-MMM-yy")
                        exSheet.Cells(3, 6) = rsCounselHist("CL_LEVEL")
                        
                        'Move back to the previous record - highest level
                        rsCounselHist.MovePrevious
                    End If
                    
                    'Retrieve Counsel History
                    Do While Not rsCounselHist.EOF
                        xCounselHist = "Level: " & rsCounselHist("CL_LEVEL") & ", Incident Date: " & Format(rsCounselHist("CL_INCDATE"), "mm/dd/yyyy") & "; " & xCounselHist
                        
                        rsCounselHist.MoveNext
                    Loop
                    exSheet.Cells(16, 2) = xCounselHist
                    
                    
                    exSheet.Cells(21, 3) = GetEmpData(rsCounsel("CL_EMPNBR"), "ED_FNAME") & " " & GetEmpData(rsCounsel("CL_EMPNBR"), "ED_SURNAME")
                    
                    'Reporting Authority
                    exSheet.Cells(23, 3) = GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU", 0), "ED_FNAME") & " " & GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU", 0), "ED_SURNAME")
                    exSheet.Cells(24, 3) = GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU2", 0), "ED_FNAME") & " " & GetEmpData(GetJHData(rsCounsel("CL_EMPNBR"), "JH_REPTAU2", 0), "ED_SURNAME")
                
                    'rsCounselHist.MoveNext
                Loop
            End If
            rsCounselHist.Close
            Set rsCounselHist = Nothing
            
            xRow = xRow + 1
Next_Emp:
            rsCounsel.MoveNext

            'exSheet.Columns.AutoFit
    
            exBook.Save
            Set exSheet = Nothing
            Set exBook = Nothing
            exApp.Quit
            Set exApp = Nothing
        Loop

        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " "
        Screen.MousePointer = DEFAULT
    
        Call Pause(1)
        If Not LanchXlsW98(xlsFileMat) Then
            Shell "cmd /c " & GetShortName(xlsFileMat)
        End If
    End If
    
    rsCounsel.Close
    Set rsCounsel = Nothing

    Screen.MousePointer = vbDefault

Exit Sub

Showa_COC_Discipline_Form_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = " "
Screen.MousePointer = DEFAULT

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Excel", "HR_COUNSEL", "Showa_COC_Discipline_Form")

Set exSheet = Nothing
Set exBook = Nothing
exApp.Quit
Set exApp = Nothing

End Sub

Private Function Get_StreamDesc(xStream) As String
    Dim xStreamDesc As String
    
    xStreamDesc = ""
    
    Select Case xStream
        Case 1: xStreamDesc = "1. Fighting with or attempting to injure another Associate."
        Case 2: xStreamDesc = "2. Possessing any weapon or related paraphernalia on SCI premises."
        Case 3: xStreamDesc = "3. Creating an intimidating, hostile, or offensive working environment by threatening, coercing, retaliating  against or using abusive/threatening language towards Associates or visitors on SCI property."
        Case 4: xStreamDesc = "4. Failing to immediately report a workplace illness, injury, incident or hazard to the Department Supervisor,  Manager or Health & Safety Specialist."
        Case 5: xStreamDesc = "5. Failing to observe established safety rules including but not limited to, not wearing appropriate PPE as per operation standards."
        Case 6: xStreamDesc = "6. Smoking or creating a flame outside the designated smoking area without authorization."
        Case 7: xStreamDesc = "7. Engaging in illegal gambling, horseplay or practical jokes, including but not limited to any prank, contest, feat of strength, unnecessary running or rough and boisterous conduct or games of chance."
        Case 8: xStreamDesc = "8. Possessing, using/misusing, distributing, selling, purchasing, reporting to work under the influence of, any illegal or prescription drugs/substances or alcohol while on SCI property."
        Case 9: xStreamDesc = "9. Stealing, willfully damaging, abusing or hiding any property belonging to another Associates or SCI."
        Case 10: xStreamDesc = "10. Giving false information with respect to absence, sickness, injury claims or personnel files."
        Case 11: xStreamDesc = "11. Misrepresenting facts or falsifying records or reports, of a business nature."
        Case 12: xStreamDesc = "12. Inappropriately obtaining or sharing confidential information with anyone outside of the appropriate SCI department."
        Case 13: xStreamDesc = "13. Interfering, failing to cooperate, or divulging confidential information related to an authorized SCI investigation."
        Case 14: xStreamDesc = "14. Obtaining property, money, or other privileges from SCI through fraud or misrepresentation or engaging in this type of activity while conducting SCI business."
        Case 15: xStreamDesc = "15. Willfully scanning or requesting the scanning of yours or another Associate’s identification card."
        Case 16: xStreamDesc = "16. Interfering with the work or performance of another associate or causing a restriction or slow down of production or tampering with or deliberately misusing company equipment."
        Case 17: xStreamDesc = "17. Using electronic equipment, cameras, video equipment, or recording devices, without proper authorization or inappropriately while on SCI premises."
        Case 18: xStreamDesc = "18. Using or carrying personal cell phones, personal paging devices or other electronic equipment on the floor unless otherwise authorized."
        Case 19: xStreamDesc = "19. Reporting to the work area without wearing the proper SCI uniform including weekends, holidays and shutdowns."
        Case 20: xStreamDesc = "20. Being absent from work for an unauthorized reason (any incident in which the associate fails to declare the absence as Personal Emergency Leave or the absence does not meet the qualifications under Personal Emergency Leave provisions)."
        Case 21: xStreamDesc = "21. Failing to contact SCI to report an absence within the scheduled shift (No call / No show)."
        Case 22: xStreamDesc = "22. Providing late notification of an absence (notification of less than 1 hour prior to shift start)."
        Case 23: xStreamDesc = "23. Failing to swipe in prior to the scheduled shift start time (late)."
        Case 24: xStreamDesc = "24. Leaving work prior to the end of the scheduled shift (Early Departure) without declaring it an Emergency Day."
        Case 25: xStreamDesc = "25. Failing to provide appropriate medical documentation following an absence of 3 or more consecutive days."
        Case 26: xStreamDesc = "26. Failing to complete the No Swipe Form and receive the Supervisor’s authorization (prior to the shift start meeting or at the end of shift). (Associates will be paid for the earliest time confirmed present at work.)"
        Case 27: xStreamDesc = "27. Reporting to the work area late and or leaving early at shift start up or after breaks and lunch or prior to shift end."
        Case 28: xStreamDesc = "28. Leaving the work area during assigned working hours without permission."
        Case 29: xStreamDesc = "29. Stopping work prior to the shift end."
        Case 30: xStreamDesc = "30. Sleeping or loafing around while on the job."
        Case 31: xStreamDesc = "31. Working below the operation standard for quality or quantity, willfully neglecting job responsibilities or refusing to comply with instructions of management."
        Case 32: xStreamDesc = "32. Removing, defacing, or changing notices or bulletins posted throughout the facility or placing signs, notes, papers, or any materials not required for production in or on products or property."
        Case 33: xStreamDesc = "33. Bringing or removing tools, materials or equipment to or from SCI premises without proper authorization."
        Case 34: xStreamDesc = "34. Duplicating any SCI keys without the proper authorization."
        Case 35: xStreamDesc = "35. Storing or posting inappropriate materials other than personal items and uniforms in SCI issued lockers."
        Case 36: xStreamDesc = "36. Parking inappropriately or outside the designated Associate parking area during working hours."
    End Select
    
    Get_StreamDesc = xStreamDesc
End Function

