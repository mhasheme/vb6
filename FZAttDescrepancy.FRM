VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRAttDiscrepancy 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Attendance/Work Schedule Discrepancy"
   ClientHeight    =   7410
   ClientLeft      =   180
   ClientTop       =   825
   ClientWidth     =   10125
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
   ScaleHeight     =   7410
   ScaleWidth      =   10125
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   5
      Left            =   2160
      TabIndex        =   9
      Top             =   3315
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   8
      Top             =   2980
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
      Left            =   2160
      TabIndex        =   7
      Top             =   2645
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
      TabIndex        =   16
      Tag             =   "Final Sorting of Records"
      Top             =   6795
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
      ItemData        =   "FZAttDescrepancy.frx":0000
      Left            =   1875
      List            =   "FZAttDescrepancy.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "First Level of grouping records"
      Top             =   6420
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   635
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
      Left            =   2160
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   300
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
      Left            =   2160
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   970
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2160
      TabIndex        =   3
      Top             =   1305
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
      Left            =   2160
      TabIndex        =   4
      Top             =   1640
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1975
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
      Left            =   2160
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2310
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
      Top             =   6480
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
   Begin Threed.SSCheck chkDescrepancy 
      Height          =   225
      Index           =   0
      Left            =   2400
      TabIndex        =   12
      Tag             =   "Show Male"
      Top             =   4320
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "  Hours > or < Work Schedule                   "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkDescrepancy 
      Height          =   225
      Index           =   1
      Left            =   2400
      TabIndex        =   13
      Tag             =   "Show Male"
      Top             =   4680
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "  No Attendance Hours That Day              "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSCheck chkDescrepancy 
      Height          =   225
      Index           =   2
      Left            =   2400
      TabIndex        =   14
      Tag             =   "Show Male"
      Top             =   5040
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "  No Work Schedule Hours That Day        "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3930
      TabIndex        =   11
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3650
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1255
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   10
      Tag             =   "40-Date from and including this date forward"
      Top             =   3650
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
      Caption         =   "Attendance From / To Date"
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
      TabIndex        =   32
      Top             =   3695
      Width           =   1965
   End
   Begin VB.Label Label2 
      Caption         =   "Type of Descrepancy:"
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
      TabIndex        =   31
      Top             =   4320
      Width           =   1635
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
      TabIndex        =   30
      Top             =   2355
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
      TabIndex        =   29
      Top             =   3360
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
      TabIndex        =   28
      Top             =   2690
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
      TabIndex        =   27
      Top             =   3025
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
      TabIndex        =   26
      Top             =   1015
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
      TabIndex        =   25
      Top             =   6825
      Width           =   825
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
      TabIndex        =   24
      Top             =   6435
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
      TabIndex        =   23
      Top             =   2020
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
      TabIndex        =   22
      Top             =   1685
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
      TabIndex        =   21
      Top             =   1350
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
      TabIndex        =   20
      Top             =   680
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
      TabIndex        =   19
      Top             =   345
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
      TabIndex        =   18
      Top             =   6120
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
      TabIndex        =   17
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmRAttDiscrepancy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQL As String
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm(frmRAttDiscrepancy.Caption & " Criteria", Me) Then Exit Sub
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
'MsgBox Error$ & " Error #:" & Str(Err)

Resume Next

End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
      Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    
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
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
      Call set_PrintState(True)
End If
Exit Sub

CRW_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "HR_ATTENDANCE", "SELECT")
Resume Next

End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub comGrpLoad()
    comGroup(0).Clear
    comGroup(2).Clear
    
    comGroup(2).AddItem "Employee Name"
    comGroup(2).AddItem "Employee Number"
    comGroup(2).Enabled = True
    comGroup(2).ListIndex = 0
    
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29, 2000
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
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
Dim SQLQ As String

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)
Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_PT
    
' dkostka - 07/05/2001 - 'Section' selection criteria was being ignored, fixed.
For x% = 0 To 5
    Call Cri_Code(x%)
Next x%
Call Cri_EE

x% = Cri_Sorts()   ' returns number of sections formated

'Make SQL script from Crystal SQL script
strSQL = ""
strSQL = Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")")

If Len(glbstrSelCri) > 0 Then
    glbstrSelCri = glbstrSelCri & " AND " & "{HRATTDESCREPWRK.AD_WRKEMP}= '" & glbUserID & "'"
End If

If comGroup(0) = "(none)" Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAttDescrep1.rpt"
Else
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAttDescrep.rpt"
End If


'Clear the records for this user from the work table
gdbAdoIhr001.BeginTrans
SQLQ = "DELETE FROM HRATTDESCREPWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans

'Call appropriate procedures to populate the temporary table for the report
If chkDescrepancy(0) Or chkDescrepancy(2) Then
    Call AttHours_WorkSchedule_Descrepancies
End If
If chkDescrepancy(1) Then
    Call NoAttendance_for_WorkScheduleDay
End If

If Len(glbstrSelCri) >= 0 Then
   Me.vbxCrystal.SelectionFormula = glbstrSelCri
End If

Me.vbxCrystal.Connect = RptODBC_SQL

' window title if appropriate
Me.vbxCrystal.WindowTitle = frmRAttDiscrepancy.Caption

Cri_SetAll = True

Screen.MousePointer = DEFAULT

Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Descrepancies", "HR_ATTENDANCE", "Select")
Cri_SetAll = False
Resume Next

End Function

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

'for labels - sort by name always
' imbeded in report

Cri_Sorts = 0
grpField$ = getEGroup(comGroup(0).Text)

If grpField$ <> "(none)" Then
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
    
    GrpIdx% = comGroup(2).ListIndex 'the item chosen in the final sort combo box
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"       'GROUP ON EMPLOYEE NAME
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = grpCond$
Else
    GrpIdx% = comGroup(2).ListIndex 'the item chosen in the final sort combo box
    Select Case GrpIdx%
        Case 0: grpField$ = "{@EFullName}"       'GROUP ON EMPLOYEE NAME
        Case 1: grpField$ = "{HREMP.ED_EMPNBR}" 'GROUP ON EMPLOYEE#
    End Select
    grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(0) = grpCond$
End If
    
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

For x% = 0 To 5
    If Not clpCode(x).ListChecker Then Exit Function
Next x%


If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Len(dlpDateRange(0).Text) > 0 Then
    If Not IsDate(dlpDateRange(0)) Then
        MsgBox "Invalid From Date"
        Me.dlpDateRange(0).SetFocus
        Exit Function
    End If
End If
If Len(dlpDateRange(1).Text) > 0 Then
    If Not IsDate(dlpDateRange(1)) Then
        MsgBox "Invalid To Date"
        Me.dlpDateRange(1).SetFocus
        Exit Function
    End If
End If
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then
        MsgBox "To Date can't be prior to From Date!"
        Me.dlpDateRange(0).SetFocus
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

glbOnTop = "FRMRATTDISCREPANCY"

Screen.MousePointer = HOURGLASS

Call setRptCaption(Me)
Call comGrpLoad

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

Private Sub AttHours_WorkSchedule_Descrepancies()
    Dim rsSchedule As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsDescrepancy As New ADODB.Recordset
    Dim SQLQ As String
    Dim xWeekDay As String
    Dim xWeekDayHrs
    
    'For each employee in the Attendance check if the Work Schedule exists
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, "
    'SQLQ = "SELECT AD_EMPNBR, AD_DOA, AD_HRS FROM HR_ATTENDANCE "
    SQLQ = SQLQ & " AD_EMPNBR, AD_DOA, AD_HRS FROM HREMP, HR_ATTENDANCE WHERE ED_EMPNBR = AD_EMPNBR "
    
    If Len(strSQL) > 0 Then
        SQLQ = SQLQ & " AND " & strSQL
    End If
    If IsDate(dlpDateRange(0).Text) Then
        SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
    End If
    If IsDate(dlpDateRange(1).Text) Then
        SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text)
    End If
    SQLQ = SQLQ & " ORDER BY AD_EMPNBR, AD_DOA ASC"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    Do While Not rsAttend.EOF
        
        'Retrieve Work Schedule that is less or equal to the Attendance Date
        SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & rsAttend("AD_EMPNBR")
        SQLQ = SQLQ & " AND SD_EDATE <= " & Date_SQL(rsAttend("AD_DOA"))
        SQLQ = SQLQ & " ORDER BY SD_EDATE DESC"         'So the most matching will be on the top
        rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsSchedule.EOF Then
            'Work Schedule(s) found for this employee
            
            rsSchedule.MoveFirst
            
            'Check which Work Schedule this Attendance record falls into
            Do While Not rsSchedule.EOF
                If rsAttend("AD_DOA") >= rsSchedule("SD_EDATE") Then
                    'Most matching Scheduler with Attendance
                    'Check if the attendance hours against the type of descrepancies
                    
                    'Get the Week Day of the Attendance
                    xWeekDay = WeekdayName(Weekday(rsAttend("AD_DOA")))
                    
                    'Get the Hours for the xWeekDay from the Work Schedule
                    xWeekDayHrs = IIf(IsNull(rsSchedule("SD_" & UCase(Left(xWeekDay, 3)) & "_HRS")), 0, rsSchedule("SD_" & UCase(Left(xWeekDay, 3)) & "_HRS"))
                    
                    'Check for Descrepancies
                    If xWeekDayHrs <> rsAttend("AD_HRS") Then
                        'Hours do not match
                        
                        If xWeekDayHrs = 0 Then
                            'No Work Schedule Hours That Day
                            If chkDescrepancy(2) Then
                                'Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_FNAME"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "No Work Schedule That Day")
                                Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), rsAttend("ED_FNAME"), rsAttend("ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "No Work Schedule That Day")
                            End If
                        ElseIf rsAttend("AD_HRS") > xWeekDayHrs Then
                            'Hours GT Work Schedule
                            If chkDescrepancy(0) Then
                                'Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_FNAME"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "Hours Greater Than Work Schedule")
                                Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), rsAttend("ED_FNAME"), rsAttend("ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "Hours Greater Than Work Schedule")
                            End If
                        ElseIf rsAttend("AD_HRS") < xWeekDayHrs Then
                            'Hours LT Work Schedule
                            If chkDescrepancy(0) Then
                                'Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_FNAME"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "Hours Less Than Work Schedule")
                                Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), rsAttend("ED_FNAME"), rsAttend("ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "Hours Less Than Work Schedule")
                            End If
                        End If
                    End If
                    
                    'Exit Do loop
                    Exit Do
                End If
                
                rsSchedule.MoveNext
            Loop
            rsSchedule.Close
            Set rsSchedule = Nothing
        Else
            'No Work Schedule Hours That Day
            If chkDescrepancy(2) Then
                'Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_FNAME"), GetEmpData(rsAttend("AD_EMPNBR"), "ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "No Work Schedule That Day")
                Call Add_Descrepancy_WorkTable(rsAttend("AD_EMPNBR"), rsAttend("ED_FNAME"), rsAttend("ED_SURNAME"), rsAttend("AD_DOA"), rsAttend("AD_HRS"), "No Work Schedule That Day")
            End If
            
            rsSchedule.Close
            Set rsSchedule = Nothing
        End If
        
Next_Attendance:
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing
    
End Sub

Private Sub NoAttendance_for_WorkScheduleDay()
    Dim rsEMP As New ADODB.Recordset
    Dim rsSchedule As New ADODB.Recordset
    Dim rsAttend As New ADODB.Recordset
    Dim rsDescrepancy As New ADODB.Recordset
    Dim SQLQ As String
    Dim recCount As Integer
    Dim recNo As Integer
    Dim xNextSchDate, xNextDate
    Dim xWeekDay As String
    Dim xWeekDayHrs
    
    'SQLQ = SQLQ & " AND SD_EDATE <= " & Date_SQL(rsAttend("AD_DOA"))
    'SQLQ = SQLQ & " WHERE AD_DOA >= " & Date_SQL(dlpDateRange(0).Text) & " AND AD_DOA <= " & Date_SQL(dlpDateRange(1).Text)
    
    'For each employee
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME FROM HREMP "
    If Len(strSQL) > 0 Then
        SQLQ = SQLQ & " WHERE " & strSQL
    End If
    SQLQ = SQLQ & " ORDER BY ED_EMPNBR ASC"
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    Do While Not rsEMP.EOF
        'Retrieve Work Schedule of the employee to check if Attendance exists
        SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & rsEMP("ED_EMPNBR")
        SQLQ = SQLQ & " ORDER BY SD_EMPNBR ASC, SD_EDATE ASC"
        rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsSchedule.EOF Then
            'Record counts
            recCount = rsSchedule.RecordCount
            recNo = 1
            
            'For each Schedule record
            Do While Not rsSchedule.EOF
                'Initialize Date field
                xNextSchDate = ""
                
                'Get the next Schedule Date
                If recCount > 1 And recNo < recCount Then
                    'More than one Work Schedule and Not the end of recordset
                    'Get next schedule's effective date - for the schedule period
                    rsSchedule.MoveNext
                    xNextSchDate = rsSchedule("SD_EDATE")
                    
                    'Move back to the previous work schedule
                    rsSchedule.MovePrevious
                End If
                
                'Check if any attendance missing for the Work Schedule period
                
                'Start Date for the loop - oldest date
                xNextDate = rsSchedule("SD_EDATE")
                
                'If Next Schedule date is not available (means it's the current schedule) then get current
                'date.
                If Not IsDate(xNextSchDate) Then xNextSchDate = Format(Now, "mm/dd/yyyy")
                
                'For each Work Schedule Day check if Attendance record is present
                Do While CVDate(xNextDate) <= CVDate(xNextSchDate)
                    'For each day of the Work Schedule
                    
                    'Get the Week Day of the Work Schedule Date
                    xWeekDay = WeekdayName(Weekday(CVDate(xNextDate)))
                    
                    'Get the Hours for the xWeekDay from the Work Schedule
                    xWeekDayHrs = IIf(IsNull(rsSchedule("SD_" & UCase(Left(xWeekDay, 3)) & "_HRS")), 0, rsSchedule("SD_" & UCase(Left(xWeekDay, 3)) & "_HRS"))
                    
                    'Only check for attendance if work schedule not 0
                    If xWeekDayHrs <> 0 Then
                        'Retrieve Attendance record for the Work Schedule day of this employee
                        SQLQ = "SELECT AD_EMPNBR, AD_DOA, AD_HRS FROM HR_ATTENDANCE "
                        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & rsSchedule("SD_EMPNBR")
                        SQLQ = SQLQ & " AND AD_DOA = " & Date_SQL(xNextDate)
                        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsAttend.EOF Then
                            'No Attendance Hours That Day
                            Call Add_Descrepancy_WorkTable(rsEMP("ED_EMPNBR"), rsEMP("ED_FNAME"), rsEMP("ED_SURNAME"), xNextDate, Null, "No Attendance Hours That Day")
                        End If
                        rsAttend.Close
                        Set rsAttend = Nothing
                    End If
                
                    'Next Day
                    xNextDate = DateAdd("d", 1, CVDate(xNextDate))
                Loop
                                
                recNo = recNo + 1
                rsSchedule.MoveNext
            Loop
            
        End If
        rsSchedule.Close
        Set rsSchedule = Nothing
        
        rsEMP.MoveNext
    Loop
    rsEMP.Close
    Set rsEMP = Nothing
End Sub

Private Sub Add_Descrepancy_WorkTable(xEMPNBR, xFName, xSurname, xDate, xHours, xReason)
    Dim rsDescrepWrk As New ADODB.Recordset
    Dim SQLQ As String
    
    'Add new record in the Descrepancy Work table
    SQLQ = "SELECT * FROM HRATTDESCREPWRK WHERE 1 = 2"
    rsDescrepWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDescrepWrk.EOF Then
        rsDescrepWrk.AddNew
        rsDescrepWrk("COMPNO") = "001"
        rsDescrepWrk("AD_EMPNBR") = xEMPNBR
        rsDescrepWrk("ED_FNAME") = xFName
        rsDescrepWrk("ED_SURNAME") = xSurname
        rsDescrepWrk("AD_DOA") = CVDate(xDate)
        rsDescrepWrk("DAYOFWEEK") = WeekdayName(Weekday(CVDate(xDate)))
        rsDescrepWrk("AD_HRS") = xHours
        rsDescrepWrk("REASON") = xReason
        rsDescrepWrk("AD_WRKEMP") = glbUserID
        rsDescrepWrk.Update
    End If
    rsDescrepWrk.Close
    Set rsDescrepWrk = Nothing
    
End Sub
