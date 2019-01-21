VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRScheduler 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Work Schedule Report"
   ClientHeight    =   6810
   ClientLeft      =   180
   ClientTop       =   825
   ClientWidth     =   9720
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
   ScaleHeight     =   6810
   ScaleWidth      =   9720
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   9
      Top             =   3645
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
      Top             =   3300
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
      Top             =   2955
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
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
      Index           =   2
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Tag             =   "Final Sorting of Records"
      Top             =   6075
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
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Tag             =   "Second level of grouping records"
      Top             =   5760
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
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "First Level of grouping records"
      Top             =   5460
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   885
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
      Top             =   540
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
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1230
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   1575
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
      Left            =   1560
      TabIndex        =   4
      Top             =   1920
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
      Top             =   2265
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
      Left            =   1560
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2610
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6840
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
   Begin INFOHR_Controls.DateLookup dlpAsOf 
      Height          =   285
      Left            =   1560
      TabIndex        =   10
      Tag             =   "40-From Date"
      Top             =   3990
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpToDate 
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Tag             =   "40-To Date"
      Top             =   3990
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   4800
      TabIndex        =   31
      Top             =   4035
      Width           =   585
   End
   Begin VB.Label lblAsOf 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Top             =   4035
      Width           =   735
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
      Left            =   120
      TabIndex        =   29
      Top             =   2655
      Width           =   1290
   End
   Begin VB.Label lblSection 
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   3690
      Width           =   540
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
      TabIndex        =   27
      Top             =   3000
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
      Top             =   3345
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
      Top             =   1275
      Width           =   615
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
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
      Index           =   2
      Left            =   120
      TabIndex        =   24
      Top             =   6135
      Width           =   825
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #2"
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
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   5820
      Width           =   885
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
      Top             =   5520
      Width           =   885
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
      TabIndex        =   21
      Top             =   2310
      Width           =   630
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
      TabIndex        =   20
      Top             =   1965
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
      TabIndex        =   19
      Top             =   1620
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
      TabIndex        =   18
      Top             =   930
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
      TabIndex        =   17
      Top             =   585
      Width           =   555
   End
   Begin VB.Label lblRepGrp 
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
      TabIndex        =   16
      Top             =   5145
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
      Left            =   0
      TabIndex        =   15
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmRScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim X%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm(frmRScheduler.Caption & " Criteria", Me) Then Exit Sub
    
    Call set_PrintState(False)
    
    X% = Cri_SetAll()
    
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

Resume Next

End Sub

Public Sub cmdView_Click()
Dim X%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
    
    Call set_PrintState(False)
      
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
      
    X% = Cri_SetAll()
    
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "HR_SCHEDULER", "SELECT")
Resume Next

End Sub

Private Sub chkSupBirthYrAge_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).Clear
    comGroup(1).Clear
    comGroup(2).Clear

  
    comGroup(1).AddItem "From Date"
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem "Employee Number"
    comGroup(1).AddItem "(none)"
    
    
    comGroup(2).AddItem "Employee Name"
    comGroup(2).AddItem "Employee Number"
    comGroup(2).Enabled = True
      
    comGroup(1).ListIndex = 0
    comGroup(2).ListIndex = 0
    
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")
    comGroup(0).AddItem lStr("Employment Status")
    comGroup(0).AddItem lStr("Category")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem lStr("Administered By")
    comGroup(0).AddItem lStr("Section")
    comGroup(0).AddItem "From Date"
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem ("Home Line")
    End If
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0

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
    Case 5: strCd$ = "HREMP.ED_SECTION" 'Lucy June 29, 2000
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

Private Sub Cri_EffectiveDate()
Dim TempCri As String
Dim TempCri1 As String
Dim dtYYY%, dtMM%, dtDD%, X%
Dim FromDate, ToDate, SQLQ

If Len(dlpAsOf.Text) = 0 And Len(dlpToDate.Text) = 0 Then Exit Sub

If Len(dlpAsOf.Text) > 0 Then
    TempCri = "({HR_SCHEDULER.SD_EDATE} "
    dtYYY% = Year(dlpAsOf.Text)
    dtMM% = month(dlpAsOf.Text)
    dtDD% = Day(dlpAsOf.Text)
    TempCri = TempCri & " >= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
End If

If Len(dlpToDate.Text) > 0 Then
    TempCri1 = "({HR_SCHEDULER.SD_TDATE} "
    dtYYY% = Year(dlpToDate.Text)
    dtMM% = month(dlpToDate.Text)
    dtDD% = Day(dlpToDate.Text)
    If Len(TempCri) > 0 Then
        TempCri1 = " AND " & TempCri1 & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    Else
        TempCri1 = TempCri1 & " <= Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    End If
End If

TempCri = TempCri & TempCri1
If Len(TempCri) > 0 Then
    GoTo Cri_FTDatst
Else
    Exit Sub
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

'Private Sub Cri_Shift()
'Dim EECri As String, OneSet%, x%
'
'If Len(txtShift.Text) < 1 Then Exit Sub
'EECri = "{HREMP.ED_SHIFT}= '" & txtShift.Text & "'"
'
'If glbiOneWhere Then
'    glbstrSelCri = glbstrSelCri & " AND " & EECri
'Else
'    glbstrSelCri = EECri
'End If
'glbiOneWhere = True
'End Sub


Private Function Cri_SetAll()
Dim X%, strRName$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
    
For X% = 0 To 5
    Call Cri_Code(X%)
Next X%
Call Cri_EE
Call Cri_EffectiveDate

X% = Cri_Sorts()   ' returns number of sections formated

Me.vbxCrystal.Formulas(10) = "showEffDate = " & IIf(IsDate(dlpAsOf) Or IsDate(dlpToDate), True, False) & ""
If IsDate(dlpAsOf) Then
    Me.vbxCrystal.Formulas(11) = "EffectiveDate=Date('" & Format(dlpAsOf.Text, "MM/dd/yyyy") & "')"
End If

If IsDate(dlpAsOf) And IsDate(dlpToDate) Then
    Me.vbxCrystal.Formulas(12) = "EffDateDesp='From Date: " & Format(dlpAsOf.Text, "MM/dd/yyyy") & " - To Date: " & Format(dlpToDate.Text, "MM/dd/yyyy") & "'"
ElseIf IsDate(dlpAsOf) Then
    'Me.vbxCrystal.Formulas(11) = "EffDateDesp=Date('" & Format(dlpAsOf.Text, "MM/dd/yyyy") & "')"
    Me.vbxCrystal.Formulas(12) = "EffDateDesp='From Date: " & Format(dlpAsOf.Text, "MM/dd/yyyy") & "'"
ElseIf IsDate(dlpToDate) Then
    'Me.vbxCrystal.Formulas(11) = "EffectiveDate=Date('" & Format(dlpToDate.Text, "MM/dd/yyyy") & "')"
    Me.vbxCrystal.Formulas(12) = "EffDateDesp='To Date: " & Format(dlpToDate.Text, "MM/dd/yyyy") & "'"
End If

If comGroup(0) = "(none)" And comGroup(1) = "(none)" Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzScheduler1.rpt"
Else
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzScheduler.rpt"
End If

If Len(glbstrSelCri) >= 0 Then
   Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL


'window title if appropriate
Me.vbxCrystal.WindowTitle = "Work Schedule Report"


Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Work Schedule", "Work Schedule Report", "Select")
Cri_SetAll = False
Resume Next

End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim X%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
'imbeded in report

Cri_Sorts = 0
grpField$ = getEGroup(comGroup(0).Text)

If grpField$ <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{HR_SCHEDULER.SD_EDATE}"
        Case 1: grpField$ = "{@EFullName}"
        Case 2: grpField$ = "{HREMP.ED_EMPNBR}"
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        If grpField$ = "{HR_SCHEDULER.SD_EDATE}" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;D"
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        End If
        Me.vbxCrystal.GroupCondition(1) = grpCond$
        
        strSFormat$ = "GH2;T;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        z% = z% + 1
        strSFormat$ = "GF2;T;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        End Select
        grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(2) = grpCond$
    Else
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        End Select
        grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    End If
Else
    dscGroup$ = "descGroup" & CStr(1) & "= 'hide'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    
    strSFormat$ = "GH1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    strSFormat$ = "GF1;F;X;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(z%) = strSFormat$
    z% = z% + 1
    
    GrpIdx% = comGroup(1).ListIndex
    Select Case GrpIdx%
        Case 0: grpField$ = "{HR_SCHEDULER.SD_EDATE}"
        Case 1: grpField$ = "{@EFullName}"
        Case 2: grpField$ = "{HREMP.ED_EMPNBR}"
        Case 3: grpField$ = "(none)"
    End Select
    If grpField$ <> "(none)" Then
        If grpField$ = "{HR_SCHEDULER.SD_EDATE}" Then
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;D"
        Else
            grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
        End If
        Me.vbxCrystal.GroupCondition(0) = grpCond$
        
        strSFormat$ = "GH2;T;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        z% = z% + 1
        strSFormat$ = "GF2;T;X;X;X;X;X;X"
        Me.vbxCrystal.SectionFormat(z%) = strSFormat$
        
        GrpIdx% = comGroup(2).ListIndex
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        End Select
        grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(1) = grpCond$
    Else
        GrpIdx% = comGroup(2).ListIndex 'the item chosen in the final sort combo box
        Select Case GrpIdx%
            Case 0: grpField$ = "{@EFullName}"
            Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
        End Select
        grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
        Me.vbxCrystal.GroupCondition(0) = grpCond$
    End If
End If

Call setRptLabel(Me, 0)

    
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

For X% = 0 To 5
    If Not clpCode(X).ListChecker Then Exit Function
Next X%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Len(dlpAsOf.Text) > 0 Then
    If Not IsDate(dlpAsOf.Text) Then
        MsgBox "Invalid From Date"
        dlpAsOf.SetFocus
        Exit Function
    End If
End If

If Len(dlpToDate.Text) > 0 Then
    If Not IsDate(dlpToDate.Text) Then
        MsgBox "Invalid To Date"
        dlpToDate.SetFocus
        Exit Function
    End If
End If

CriCheck = True
End Function

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = "FRMRSCHEDULER"

Screen.MousePointer = HOURGLASS

Call comGrpLoad
Call setRptCaption(Me)

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6

If glbLinamar Then clpCode(3).MaxLength = 8

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

Private Sub txtMonth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtMonth_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
End Sub

Private Sub txtShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtYear_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then KeyAscii = 0
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

