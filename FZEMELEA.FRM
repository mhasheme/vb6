VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmREmergLeave 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Emergency Leave Report"
   ClientHeight    =   7920
   ClientLeft      =   360
   ClientTop       =   1005
   ClientWidth     =   10410
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7920
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkInclAtt 
      Caption         =   "Include Attendance History"
      Height          =   375
      Left            =   5460
      TabIndex        =   17
      Top             =   5115
      Width           =   2355
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2595
      MaxLength       =   4
      TabIndex        =   13
      Tag             =   "00-Shift"
      Top             =   4284
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtEmerLea 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2595
      MaxLength       =   7
      TabIndex        =   14
      Tag             =   "10-Emergency Leave (in days)"
      Top             =   4620
      Width           =   900
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "Final sort of records"
      Top             =   7095
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   2
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "Third level of grouping records"
      Top             =   6780
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   1
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "Second level of grouping records"
      Top             =   6465
      Width           =   2325
   End
   Begin VB.ComboBox comGroup 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   0
      Left            =   2130
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Tag             =   "First Level of grouping records"
      Top             =   6150
      Width           =   2325
   End
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   0
      Left            =   2325
      TabIndex        =   15
      Tag             =   "Detailed Report"
      Top             =   5175
      Width           =   1470
      _Version        =   65536
      _ExtentX        =   2593
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Detailed"
      ForeColor       =   16711680
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
   Begin Threed.SSOption optGrouping 
      Height          =   255
      Index           =   1
      Left            =   3780
      TabIndex        =   16
      Tag             =   "Summary Report"
      Top             =   5175
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  Summary"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   1668
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
      Left            =   2280
      TabIndex        =   5
      Tag             =   "EDPT-Category"
      Top             =   1995
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
      Left            =   2280
      TabIndex        =   3
      Tag             =   "00-Enter Union Code"
      Top             =   1341
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
      Left            =   2280
      TabIndex        =   2
      Tag             =   "00-Enter Location Code"
      Top             =   1014
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   687
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
      Left            =   2280
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
      Left            =   2280
      TabIndex        =   11
      Tag             =   "00-Enter Administered By Code"
      Top             =   3630
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
      Left            =   2280
      TabIndex        =   12
      Tag             =   "00-Enter Section Code"
      Top             =   3957
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   10
      Tag             =   "00-Enter Region Code"
      Top             =   3303
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   4140
      TabIndex        =   9
      Tag             =   "40-Date upto and including this date forward"
      Top             =   2976
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   8
      Tag             =   "40-Date from and including this date forward"
      Top             =   2976
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   2322
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptau 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Tag             =   "10-Reporting Authority"
      Top             =   2649
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
      Left            =   0
      Top             =   7485
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   42
      Top             =   4329
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   41
      Top             =   2040
      Width           =   630
   End
   Begin VB.Label lblEmer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Emergency Leave (in days)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   40
      Top             =   4665
      Width           =   2310
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   39
      Top             =   3675
      Width           =   1125
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   38
      Top             =   4002
      Width           =   540
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   37
      Top             =   3348
      Width           =   510
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   36
      Top             =   1059
      Width           =   615
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Final Sort"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   210
      TabIndex        =   35
      Top             =   7125
      Width           =   660
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #3"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   210
      TabIndex        =   34
      Top             =   6810
      Width           =   885
   End
   Begin VB.Label lblGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Grouping #2"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   33
      Top             =   6495
      Width           =   885
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
      Left            =   210
      TabIndex        =   32
      Top             =   6180
      Width           =   885
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
      TabIndex        =   31
      Top             =   5820
      Width           =   1575
   End
   Begin VB.Label lblTypeRep 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Report"
      Height          =   195
      Left            =   180
      TabIndex        =   30
      Top             =   5205
      Width           =   1065
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   29
      Top             =   3021
      Width           =   1320
   End
   Begin VB.Label lblRepAuth 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   28
      Top             =   2694
      Width           =   1350
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   27
      Top             =   2367
      Width           =   1290
   End
   Begin VB.Label lblCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   26
      Top             =   1713
      Width           =   450
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   25
      Top             =   1386
      Width           =   420
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   24
      Top             =   732
      Width           =   825
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   23
      Top             =   405
      Width           =   555
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
      TabIndex        =   22
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmREmergLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim HisSQL
Dim WRKSQL
Dim flgUseLabelMst As Boolean

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Emergency Leave Report Criteria", Me) Then Exit Sub
    'cmdPrint.Enabled = False
    'cmdView.Enabled = False
     Call set_PrintState(False)
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
    'Me.vbxCrystal.Password = gstrAccPWord$
    'Me.vbxCrystal.UserName = gstrAccUID$
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    Call set_PrintState(True)
'    cmdPrint.Enabled = True
'    cmdView.Enabled = True
End If
Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Public Sub cmdView_Click()
Dim x%
Dim strWHand As String
On Error GoTo CRW_Err

If CriCheck() Then
    Screen.MousePointer = HOURGLASS
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    Call set_PrintState(False)
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
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString
Resume Next
Screen.MousePointer = DEFAULT
End Sub

Private Sub comGroup_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comGrpLoad()
Dim x%

    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Location")      'Jaddy jun 16, 1999
    comGroup(0).AddItem "Employee Name" 'laura nov 24, 1997
    comGroup(0).AddItem lStr("G/L")
    comGroup(0).AddItem lStr("Section")  'Lucy June 29,2000
    
    If glbLinamar Then ' Frank May 2,2001
        comGroup(0).AddItem "Employment Type"
        comGroup(0).AddItem ("Home Line")
    End If
    If Not glbMulti Then comGroup(0).AddItem "Shift"
    comGroup(0).AddItem lStr("Region")
    comGroup(0).AddItem lStr("AttSupervisor")
    comGroup(0).AddItem "(none)"
    comGroup(0).ListIndex = 0
    
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem lStr("Union")
    comGroup(1).AddItem "(none)"
    comGroup(1).ListIndex = 0
    
    comGroup(2).AddItem "Employee Name"
    comGroup(2).AddItem "(none)"
    comGroup(2).ListIndex = 0
    comGroup(3).AddItem "Attendance Date"
    comGroup(3).ListIndex = 0

End Sub

Private Sub Cri_Code(intIdx%)
Dim CodeCri As String, fld$
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim strCd$

If Len(clpCode(intIdx%).Text) > 0 Then
    Select Case intIdx%
    Case 0: strCd$ = "HREMP.ED_LOC"
    Case 1: strCd$ = "HREMP.ED_ORG" 'IIf(glbMulti, "HR_ATTENDANCE.AD_ORG", "HREMP.ED_ORG")
    Case 2: strCd$ = "HREMP.ED_EMP"
    Case 3: strCd$ = "HREMP.ED_REGION"
    Case 4: strCd$ = "HREMP.ED_ADMINBY"
    Case 5: strCd$ = "HREMP.ED_SECTION"
    'Case 6: strCd$ = "HRJOB.JB_GRPCD"
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

Private Sub Cri_Sup()
Dim EECri As String

If Len(elpReptau.Text) > 0 Then
    EECri = "{HR_JOB_HISTORY.JH_REPTAU} IN [" & getEmpnbr(elpReptau.Text) & "] "
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

Private Sub Cri_FTDates()

Dim TempCri As String
Dim TempCri01 As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    HisSQL = HisSQL & "AND ((HR_ATTENDANCE.AD_DOA) Between "
    HisSQL = HisSQL & Date_SQL(dlpDateRange(0).Text) & " And "
    HisSQL = HisSQL & Date_SQL(dlpDateRange(1).Text) & ") "
    TempCri = "({HRATTWRK.AD_DOA} "
    TempCri01 = "{HRATTWRK.AD_DOA} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    TempCri01 = "(" & TempCri01 & " >= " & Date_SQL(dlpDateRange(0).Text) & " AND " & TempCri01 & " <= " & Date_SQL(dlpDateRange(1).Text) & ") "
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(x%).Text) > 0 Then
        TempCri = "({HRATTWRK.AD_DOA}  "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(x%).Text)
        dtMM% = month(dlpDateRange(x%).Text)
        dtDD% = Day(dlpDateRange(x%).Text)
        TempCri01 = TempCri & Date_SQL(dlpDateRange(x%).Text) & ")"
        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
        If x% = 0 Then
            HisSQL = HisSQL & "AND HR_ATTENDANCE.AD_DOA >= " & Date_SQL(dlpDateRange(0).Text)
        Else
            HisSQL = HisSQL & "AND HR_ATTENDANCE.AD_DOA <= " & Date_SQL(dlpDateRange(1).Text)
        End If
        GoTo Cri_FTDatst
    End If
Next x%



Cri_FTDatst:
If Len(TempCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = TempCri
        WRKSQL = TempCri01
    Else
        glbstrSelCri = glbstrSelCri & " AND " & TempCri
        WRKSQL = WRKSQL & " AND " & TempCri01
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
Dim x%, xNoFile
Dim strSFormat$

Cri_SetAll = False

On Error GoTo modSetCriteria_Err

Screen.MousePointer = HOURGLASS

flgUseLabelMst = True
glbiOneWhere = False
glbstrSelCri = ""
WRKSQL = ""

' call cri models set both glbiONeWhere and strSelCri
'Call glbCri_Dept(Me)  'laura nov 22, 1997
Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_EE
Call Cri_PT
Call Cri_Shift
For x% = 0 To 5 '6 '4
    Call Cri_Code(x%)
Next x%


HisSQL = " AD_EMELEA <> 0 AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & Replace(Replace(Replace(Replace(glbstrSelCri, "{", ""), "}", ""), "[", "("), "]", ")") & ")"

WRKSQL = glbstrSelCri
Call Cri_FTDates
WRKSQL = Replace(WRKSQL, "{", "(")
WRKSQL = Replace(WRKSQL, "}", ")")

'If Not (glbSQL And optGrouping(1)) Then
    Call SELATTWRK
'End If
Call Cri_Sup
If optGrouping(0) Then
    If Not glbLinamar Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzemeled.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rlemeled.rpt"
    End If
Else
    flgUseLabelMst = False
    If Not glbLinamar Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzemeles.rpt"
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rlemeley.rpt"
    End If
End If

'If Not (glbLinamar And optGrouping(1)) Then
    Me.vbxCrystal.SelectionFormula = Replace(glbstrSelCri, "HR_ATTENDANCE", "HRATTWRK") & " AND ({HRATTWRK.AD_WRKEMP} = '" & glbUserID & "' )"
'Else
'    Me.vbxCrystal.SelectionFormula = glbstrSelCri
'End If

x% = Cri_Sorts()
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
    Me.vbxCrystal.DataFiles(x%) = glbIHRDBW
    For x% = 1 To 8
        Me.vbxCrystal.DataFiles(x%) = glbIHRDB
    Next x%
    
End If
Me.vbxCrystal.WindowTitle = "Emergency Leave Report"
Cri_SetAll = True
Screen.MousePointer = DEFAULT
Exit Function


modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Emergency Leave", "Emergency Leave Report", "Select")
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
EECri = "{HR_ATTENDANCE.AD_SHIFT}= '" & txtShift.Text & "'"

If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True
End Sub

Private Function Cri_Sorts()
Dim grpCond$, grpField$
Dim x%, Y%, z%, strSFormat$, strVis$, strFVis$, strPage$
Dim dscGroup$, GrpIdx%

Cri_Sorts = 0

'Use Label Master
If flgUseLabelMst = True Then
    Call setRptLabel(Me, 4)
End If

' first set primary grouping
If comGroup(0).Text = lStr("AttSupervisor") Then
    grpField$ = "{@fldADSuper}"
Else
    grpField$ = getEGroup(comGroup(0).Text)
    If grpField$ = "(none)" Then grpField$ = "{HREMP.ED_COMPNO}"
End If

grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$

If Not (grpField$ = "{HREMP.ED_COMPNO}") Then 'GrpIdx% <> 5 Then     ' IF Not (none) - SHOW GROUP1  'laura nov 24, 1997 changed from 2 to 3
    dscGroup$ = comGroup(0).Text
    dscGroup$ = "descGroup1 = '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descGroup2 = " & grpField$
    Me.vbxCrystal.Formulas(1) = dscGroup$
    Me.vbxCrystal.SectionFormat(1) = "GF3;T;F;F;X;X;X;X"
    Me.vbxCrystal.SectionFormat(0) = "GH1;T;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(2) = "GF1;T;X;F;X;X;X;X"
Else
    dscGroup$ = "descGroup1 = ''"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descGroup2 = ''"
    Me.vbxCrystal.Formulas(1) = dscGroup$
    Me.vbxCrystal.SectionFormat(0) = "GH1;F;T;X;X;X;X;X"
    Me.vbxCrystal.SectionFormat(2) = "GF1;F;X;F;X;X;X;X"
End If
Me.vbxCrystal.Formulas(3) = "EmeEarned = " & txtEmerLea

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{HREMP.ED_ORG}"
    Case 2: grpField$ = "{@EFullName}"    '(none)
End Select

grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(1) = grpCond$

GrpIdx% = comGroup(2).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{@EFullName}"
    Case 1: grpField$ = "{@EFullName}"    '(none)
End Select
        
grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(2) = grpCond$

If comGroup(1).ListIndex > 0 Then
    Me.vbxCrystal.SectionFormat(0) = "GF1;F;X;X;X;X;X;X"
End If

If dlpDateRange(0).Text <> "" And dlpDateRange(1).Text <> "" Then
    strSFormat$ = "As of " & dlpDateRange(0).Text & " through " & dlpDateRange(1).Text
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"

Else
    strSFormat$ = "No date entered"
    Me.vbxCrystal.Formulas(2) = "Daterange = '" & strSFormat$ & "'"

End If
Cri_Sorts = z%
End Function

Private Function CriCheck()
Dim I%

CriCheck = False

For I% = 0 To 5 '6 '5
    If Not clpCode(I).ListChecker Then Exit Function

Next I%

If Not clpDept.ListChecker Then
'If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    'MsgBox "If Department Entered - it must be known"
    'clpDept.SetFocus
    Exit Function
End If

If Not clpDiv.ListChecker Then
'If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then      'Serbo
    'MsgBox "If Team Entered - it must be known"                     '
    'clpDiv.SetFocus                                                '
    Exit Function                                                   '
End If

'If Len(txtEEID) > 0 And lblEEName.Caption = "Enter Valid Employee #" Then
'    MsgBox "If Employee Entered - they must exist"
'    txtEEID.SetFocus
'    Exit Function
'End If


For I% = 0 To 1
    If Len(dlpDateRange(I%).Text) > 0 Then
       If Not IsDate(dlpDateRange(I%)) Then
           MsgBox "Not a valid date"
           dlpDateRange(I%).SetFocus
           Exit Function
       End If
    Else
       'Jaddy changed on Jan 8, 2003 by Jerry's request.
       MsgBox "Date Range are required fields."
       dlpDateRange(I%).SetFocus
       Exit Function
    End If
Next I%

If Not clpPT.ListChecker Then
'If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    'MsgBox lStr("Category code must be valid")
    'clpPT.SetFocus
    Exit Function
End If

If Len(txtEmerLea) <= 0 Then
    MsgBox "Emergency Leave is required"
    txtEmerLea.SetFocus
    Exit Function
Else
    If Len(txtEmerLea) > 0 And Not IsNumeric(txtEmerLea) Then
        MsgBox "If Emergency Leave Entered - it must be number"
        txtEmerLea.SetFocus
        Exit Function
    End If
End If

If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
    End If
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

glbOnTop = "FRMREMERGLEAVE"

Screen.MousePointer = HOURGLASS

Call comGrpLoad

Call setRptCaption(Me)
Call setCaption(lblShift)
lblFromTo.Caption = lStr("From Date") & " / " & lStr("To Date")

If glbCompSerial = "S/N - 2227W" Then clpCode(3).MaxLength = 6
If glbCompSerial = "S/N - 2381W" Or glbCompSerial = "S/N - 2411W" Then clpCode(0).MaxLength = 6
If glbMulti Then lblUnion.ForeColor = &HC000C0

If glbLinamar Then
    clpCode(3).MaxLength = 8
    txtEmerLea = 10
    txtEmerLea.Enabled = False
End If
Call INI_Controls(Me)
If Not glbMulti Then
    lblShift.Visible = True
    txtShift.Visible = True
End If
If Not gSec_Inq_Attendance_History Then
    chkInclAtt.Enabled = False
End If

Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False


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

Private Sub optGrouping_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub


Private Sub SELATTWRK()
Dim xlen, xxx, xx1
Dim db001 As Database
Dim RsTempDAO As Recordset
Dim RsTempADO As New ADODB.Recordset
Dim SQLQ, OEmpNbr
On Error GoTo AttWrkError
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(0).FloodPercent = 1
MDIMain.panHelp(0).FloodPercent = 3
gdbAdoIhr001.CommandTimeout = 600
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).FloodPercent = 15
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute "DELETE FROM HRATTWRK " & in_SQL(glbIHRDBW) & " WHERE AD_WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.CommitTrans

MDIMain.panHelp(0).FloodPercent = 30
Dim xField As String
xField = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
SQLQ = "INSERT INTO HRATTWRK (" & xField & ",AD_WRKEMP)"
SQLQ = SQLQ & in_SQL(glbIHRDBW)
SQLQ = SQLQ & "SELECT " & xField
SQLQ = SQLQ & ",'" & glbUserID & "' AS AD_WRKEMP "
SQLQ = SQLQ & " FROM HR_ATTENDANCE "
If Len(HisSQL) > 1 Then
    SQLQ = SQLQ & " WHERE (" & HisSQL & ")"
Else
    SQLQ = SQLQ
End If

MDIMain.panHelp(0).FloodPercent = 45
gdbAdoIhr001.BeginTrans
gdbAdoIhr001.Execute SQLQ
gdbAdoIhr001.CommitTrans
MDIMain.panHelp(0).FloodPercent = 55

If chkInclAtt Then
    SQLQ = "INSERT INTO HRATTWRK (AD_COMPNO, AD_EMPNBR, AD_DOA, "
    SQLQ = SQLQ & "AD_REASON_TABL, AD_REASON, AD_HRS, AD_COMM, "
    SQLQ = SQLQ & "AD_CHRGCODE, AD_SHIFT, AD_SUPER, AD_INCID, AD_SEN, "
    SQLQ = SQLQ & "AD_LDATE, AD_LTIME, AD_LUSER , AD_WRKEMP) "
    SQLQ = SQLQ & in_SQL(glbIHRDBW)
    SQLQ = SQLQ & "SELECT AH_COMPNO, "
    SQLQ = SQLQ & "AH_EMPNBR, AH_DOA, "
    SQLQ = SQLQ & "AH_REASON_TABL, AH_REASON, "
    SQLQ = SQLQ & "AH_HRS, AH_COMM, "
    SQLQ = SQLQ & "AH_CHRGCODE, AH_SHIFT, "
    SQLQ = SQLQ & "AH_SUPER, AH_INCID, "
    SQLQ = SQLQ & "AH_SEN, AH_LDATE, "
    SQLQ = SQLQ & "AH_LTIME, AH_LUSER, '" & glbUserID & "' " 'AH_LUSER "
    SQLQ = SQLQ & "FROM HR_ATTENDANCE_HISTORY "
    If Len(HisSQL) > 1 Then
        SQLQ = SQLQ & "WHERE (" & Replace(Replace(HisSQL, "AD_", "AH_"), "HR_ATTENDANCE", "HR_ATTENDANCE_HISTORY") & ")"
    Else
        SQLQ = SQLQ & ""
    End If
    MDIMain.panHelp(0).FloodPercent = 75
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    If Not glbSQL And Not glbOracle Then Call Pause(1)
End If

'Comment by Frank ticket# 7364
'If glbCompSerial = "S/N - 2350W" Then 'Ticket# 7258 Listowel Tech
'    SQLQ = "SELECT * FROM HRATTWRK WHERE AD_WRKEMP = '" & glbUserID & "' "
'    RsTempADO.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
'    Do While Not RsTempADO.EOF
'        RsTempADO("AD_HRS") = 0
'        If Not IsNull(RsTempADO("AD_CHRGCODE")) Then
'            If IsNumeric(RsTempADO("AD_CHRGCODE")) Then
'                RsTempADO("AD_HRS") = RsTempADO("AD_CHRGCODE")
'            End If
'        End If
'        RsTempADO.MoveNext
'    Loop
'    RsTempADO.Close
'    SQLQ = "DELETE FROM HRATTWRK WHERE AD_HRS =0 AND AD_WRKEMP = '" & glbUserID & "' "
'    gdbAdoIhr001W.Execute SQLQ
'
'End If

MDIMain.panHelp(0).FloodPercent = 60

HisSQL = ""

SQLQ = "SELECT * FROM HRATTWRK "
SQLQ = SQLQ & "WHERE AD_WRKEMP = '" & glbUserID & "' "
SQLQ = SQLQ & "ORDER BY AD_EMPNBR  "
RsTempADO.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
OEmpNbr = -1
Do While Not RsTempADO.EOF
    If RsTempADO("AD_EMPNBR") <> OEmpNbr Then
        RsTempADO("AD_WHRS") = Val(txtEmerLea)
        RsTempADO.Update
        OEmpNbr = RsTempADO("AD_EMPNBR")
    Else
        RsTempADO("AD_WHRS") = 0
        RsTempADO.Update
    End If
    RsTempADO.MoveNext
Loop
RsTempADO.Close

gdbAdoIhr001.CommandTimeout = 600
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

AttWrkError:
    gdbAdoIhr001.CommandTimeout = 600
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Exit Sub
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


