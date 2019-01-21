VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmRPERFORMANCEREVIEW 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Staff Profile Report"
   ClientHeight    =   8130
   ClientLeft      =   570
   ClientTop       =   675
   ClientWidth     =   10785
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
   ScaleHeight     =   8130
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
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
      Index           =   0
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Tag             =   "First Level of grouping records"
      Top             =   4515
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Tag             =   "Second level of grouping records"
      Top             =   4845
      Width           =   2325
   End
   Begin INFOHR_Controls.CodeLookup clpEvent 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "SDPE"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "Performance Category"
      Top             =   1680
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "SDPG"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "00-Specific Department Desired"
      Top             =   960
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
      Left            =   1920
      TabIndex        =   0
      Tag             =   "00-Specific Division Desired"
      Top             =   600
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
   Begin INFOHR_Controls.CodeLookup clpPosition 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "00-Enter Position Group  Code"
      Top             =   1320
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "JBGC"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      Height          =   285
      Index           =   1
      Left            =   3750
      TabIndex        =   9
      Tag             =   "40-Date upto and including this date forward"
      Top             =   3480
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
      TabIndex        =   8
      Tag             =   "40-Date from and including this date forward"
      Top             =   3480
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   22
      Top             =   7710
      Width           =   10785
      _Version        =   65536
      _ExtentX        =   19024
      _ExtentY        =   741
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   60
         Top             =   30
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
   End
   Begin INFOHR_Controls.EmployeeLookup elpRept 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Tag             =   "00-Reporting Authority 1"
      Top             =   3120
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   503
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Tag             =   "00-Enter Position Code "
      Top             =   2760
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin VB.Label lblPosition 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Code"
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
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority"
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
      Top             =   3120
      Width           =   1350
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee #"
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
      Top             =   2400
      Width           =   840
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
      Left            =   60
      TabIndex        =   21
      Top             =   4200
      Width           =   1575
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
      TabIndex        =   20
      Top             =   4575
      Width           =   885
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
      TabIndex        =   19
      Top             =   4905
      Width           =   885
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Performance Category"
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
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label lblPosGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
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
      Top             =   1320
      Width           =   1035
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
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblBCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Performance  Event"
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
      Top             =   2040
      Width           =   1410
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
      TabIndex        =   14
      Top             =   3510
      Width           =   1095
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
      TabIndex        =   13
      Top             =   960
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
      TabIndex        =   12
      Top             =   600
      Width           =   555
   End
End
Attribute VB_Name = "frmRPERFORMANCEREVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEmpList
Dim FNotMan, FYesMan

Private Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr

If CriCheck() Then
    If Not PrtForm("Staff Profile Report", Me) Then Exit Sub
    Call set_PrintState(False)
'    cmdPrint.Enabled = False
'    cmdView.Enabled = False
    x% = Cri_SetAll()
    Me.vbxCrystal.Destination = 1
    MDIMain.Timer1.Enabled = False
   
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    Me.vbxCrystal.Reset
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CRW", "SALPRF", "SELECT")
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
Dim vPosGroup As String
    If Not glbSyndesis Then
        vPosGroup = "Position Group"
    Else
        vPosGroup = "Position Grade"
    End If
    comGroup(0).AddItem lStr("Division")
    comGroup(0).AddItem lStr("Department")
    comGroup(0).AddItem lStr("Performance Category")
    comGroup(0).AddItem lStr("Performance Event")
    comGroup(0).AddItem "Employee Name"
    comGroup(0).AddItem vPosGroup
    comGroup(0).AddItem lStr("Position Code")
    comGroup(0).AddItem "Shift"
    comGroup(0).AddItem "(none)"
    
    comGroup(1).AddItem lStr("Division")
    comGroup(1).AddItem lStr("Department")
    comGroup(1).AddItem lStr("Performance Category")
    comGroup(1).AddItem lStr("Performance Event")
    comGroup(1).AddItem "Employee Name"
    comGroup(1).AddItem vPosGroup
    comGroup(1).AddItem lStr("Position Code")
    comGroup(1).AddItem "Shift"
    comGroup(1).AddItem "(none)"
    
    comGroup(0).ListIndex = 0
    comGroup(1).ListIndex = 0
    
  
  
End Sub

Private Sub Cri_Position()
Dim PosCri As String


If Len(clpPosition.Text) > 0 Then
    PosCri = "({HRJOB.JB_GRPCD} IN ['" & Replace(clpPosition.Text, ",", "','") & "'])"
End If

If Len(PosCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = PosCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & PosCri
    End If
    glbiOneWhere = True
End If

End Sub
Private Sub Cri_Event()
Dim EventCri As String


If Len(clpEvent.Text) > 0 Then
    EventCri = "({HR_PERFORM_FRIESEN.PH_EVENTCODE} IN ['" & Replace(clpEvent.Text, ",", "','") & "'])"
End If

If Len(EventCri) >= 1 Then
    If Not glbiOneWhere Then
        glbstrSelCri = EventCri
    Else
        glbstrSelCri = glbstrSelCri & " AND " & EventCri
    End If
    glbiOneWhere = True
End If

End Sub
Private Sub Cri_Div()

Dim DivCri As String


If Len(clpDiv.Text) > 0 Then
    DivCri = "({HREMP.ED_DIV} IN ['" & Replace(clpDiv.Text, ",", "','") & "'])"
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


Private Sub Cri_FTDates()

Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%, strCriNam$
 strCriNam$ = "HR_PERFORM_FRIESEN.PH_PREVIEW"
If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
    TempCri = "({" & strCriNam$ & "} "
    dtYYY% = Year(dlpDateRange(0).Text)
    dtMM% = month(dlpDateRange(0).Text)
    dtDD% = Day(dlpDateRange(0).Text)
    TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
    dtYYY% = Year(dlpDateRange(1).Text)
    dtMM% = month(dlpDateRange(1).Text)
    dtDD% = Day(dlpDateRange(1).Text)
    TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
    GoTo Cri_FTDatst
End If

For x% = 0 To 1
    If Len(dlpDateRange(0).Text) > 0 Then
        TempCri = "({" & strCriNam$ & "} "
        If x% = 0 Then
            TempCri = TempCri & " >= "
        Else
            TempCri = TempCri & " <= "
        End If
        dtYYY% = Year(dlpDateRange(0).Text)
        dtMM% = month(dlpDateRange(0).Text)
        dtDD% = Day(dlpDateRange(0).Text)
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
   
   
   
   
'If glbiOneWhere Then
'    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
'
'        glbstrSelCri = glbstrSelCri & " AND {HR_PERFORM_FRIESEN.PH_PREVIEW} >= " & dlpDateRange(0) & " AND {HR_PERFORM_FRIESEN.PH_PREVIEW} <= " & dlpDateRange(1)
'    ElseIf Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) > 0 Then
'           glbstrSelCri = glbstrSelCri & " AND {HR_PERFORM_FRIESEN.PH_PREVIEW} <= " & dlpDateRange(1)
'    ElseIf Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) = 0 Then
'           glbstrSelCri = glbstrSelCri & " AND {HR_PERFORM_FRIESEN.PH_PREVIEW} >= " & dlpDateRange(0)
'    End If
'Else
'    If Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) > 0 Then
'       glbstrSelCri = " {HR_PERFORM_FRIESEN.PH_PREVIEW} >= " & dlpDateRange(0) & " AND {HR_PERFORM_FRIESEN.PH_PREVIEW} <= " & dlpDateRange(2)
'    ElseIf Len(dlpDateRange(0).Text) = 0 And Len(dlpDateRange(1).Text) > 0 Then
'           glbstrSelCri = " {HR_PERFORM_FRIESEN.PH_PREVIEW} <= " & dlpDateRange(1)
'    ElseIf Len(dlpDateRange(0).Text) > 0 And Len(dlpDateRange(1).Text) = 0 Then
'           glbstrSelCri = " {HR_PERFORM_FRIESEN.PH_PREVIEW} >= " & dlpDateRange(0)
'    End If
'End If
'        glbiOneWhere = True
  
   
'    dtYYY% = Year(dlpDateRange(0).Text)
'    dtMM% = Month(dlpDateRange(0).Text)
'    dtDD% = Day(dlpDateRange(0).Text)
'    TempCri = TempCri & " in [Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")] "
'    dtYYY% = Year(dlpDateRange(1).Text)
'    dtMM% = Month(dlpDateRange(1).Text)
'    dtDD% = Day(dlpDateRange(1).Text)
'    TempCri = TempCri & " to [Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")] "
'    GoTo Cri_FTDatst
'End If

'For x% = 0 To 1
'    If Len(dlpDateRange(0).Text) > 0 Then
'        TempCri = "({" & strCriNam$ & "} "
'        If x% = 0 Then
'            TempCri = TempCri & " >= "
'        Else
'            TempCri = TempCri & " <= "
'        End If
'        dtYYY% = Year(dlpDateRange(0).Text)
'        dtMM% = Month(dlpDateRange(0).Text)
'        dtDD% = Day(dlpDateRange(0).Text)
'        TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
'        GoTo Cri_FTDatst
'    End If
'Next x%



''Cri_FTDatst:
'If Len(TempCri) >= 1 Then
'    If Not glbiOneWhere Then
'        glbstrSelCri = TempCri
'    Else
'        glbstrSelCri = glbstrSelCri & " AND " & TempCri
'    End If
'    glbiOneWhere = True
'End If

End Sub

Private Sub Cri_Category()
Dim EECri As String

If Len(clpPT.Text) < 1 Then Exit Sub

EECri = "({HR_PERFORM_FRIESEN.PH_CATECODE} IN ['" & Replace(clpPT.Text, ",", "','") & "'])"


If glbiOneWhere Then
    glbstrSelCri = glbstrSelCri & " AND " & EECri
Else
    glbstrSelCri = EECri
End If
glbiOneWhere = True


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
Private Sub Cri_PositionCode()
Dim countr   As Integer  ' EEList_Snap is definded at form level
Dim PosCri As String
If Len(clpJob.Text) <= 0 Then Exit Sub
PosCri = "({HR_JOB_HISTORY.JH_JOB} = '" & clpJob.Text & "')"
If Len(glbstrSelCri) > 1 Then
    glbstrSelCri = glbstrSelCri & " AND " & PosCri
Else
    glbstrSelCri = PosCri
End If

End Sub
Private Sub Cri_RepAuth()
Dim EECri As String

If Len(elpRept.Text) > 0 Then
    EECri = "{HR_PERFORM_FRIESEN.PH_REPTAU}= " & Trim(elpRept.Text) & " "
   
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
Dim x%, XFMT, LocstrSelCri

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""

Call glbCri_DeptUN(clpDept.Text)

Call Cri_Div    ' sets fglbCriteria and fglbiOneWhere
Call Cri_Category

Call Cri_Position
Call Cri_Event
Call Cri_PositionCode

Call Cri_FTDates
Call Cri_EE
Call Cri_RepAuth
' report name


Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzprfrvu.rpt"
    
x% = Cri_Sorts()
LocstrSelCri = ""
        
'Ticket #28635 - Add View Own security
If Len(glbUserEmpNo) > 0 And glbUserEmpNo <> 0 And Not gSec_Performance_ViewOwn Then
    'Do not show user's Performance records based on the Employee # associated to the User.
    If Len(glbstrSelCri) > 0 Then
        glbstrSelCri = glbstrSelCri & " AND {HR_PERFORM_FRIESEN.PH_EMPNBR} <> " & glbUserEmpNo
    Else
        glbstrSelCri = glbstrSelCri & " {HR_PERFORM_FRIESEN.PH_EMPNBR} <> " & glbUserEmpNo
    End If
End If
        
If Len(glbstrSelCri) >= 0 Then
    LocstrSelCri = glbstrSelCri
End If
    
    
Me.vbxCrystal.SelectionFormula = LocstrSelCri

Me.vbxCrystal.Connect = RptODBC_SQL
 

Me.vbxCrystal.WindowTitle = "Staff Profile Report"

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
Dim dscGroup$, GrpIdx%, SavGrp1, SavGrp2, SavGrp3, SavGrp4
Cri_Sorts = 0

' first set primary grouping
grpField$ = getEGroup(comGroup(0).Text)
If comGroup(0).Text = "(none)" Then grpField$ = "{HRPARCO.PC_NAME}"
If comGroup(0) <> "(none)" Then
    'If comGroup(0).Text = "Performance Catergory" Then
        dscGroup$ = comGroup(0).Text
    
    dscGroup$ = "descGroup" & CStr(1) & "= '" & dscGroup$ & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    dscGroup$ = "descName1 = " & grpField$ 'added by RAUBREY 5/15/97
    Me.vbxCrystal.Formulas(1) = dscGroup$ 'added by RAUBREY 5/15/97
Else
    Me.vbxCrystal.Formulas(0) = "descGroup1= ''"
    Me.vbxCrystal.Formulas(1) = "descName1 = ''"
End If

grpCond$ = "GROUP" & CStr(1) & ";" & grpField$ & ";ANYCHANGE;A"
Me.vbxCrystal.GroupCondition(0) = grpCond$

GrpIdx% = comGroup(1).ListIndex
Select Case GrpIdx%
    Case 0: grpField$ = "{HRDEPT.DF_NAME}"
    Case 1: grpField$ = "{HR_DIVISION.Division_Name}"
    Case 2: grpField$ = "{tblCategory.TB_DESC}"
    Case 3: grpField$ = "{tblEvent.TB_DESC}"
    Case 4: grpField$ = "{@EFullName}"
    Case 5: grpField$ = "{tblPosGroup.TB_DESC}"
    Case 6: grpField$ = "{HRJOB.JB_DESCR}"
    Case 7: grpField$ = "{HREMP.ED_SHIFT}"
    Case 8: grpField$ = "(none)"
End Select
If grpField$ = "(none)" Then
grpField$ = "{HR_PERFORM_FRIESEN.PH_PREVIEW}"
Else
    grpCond$ = "GROUP" & CStr(2) & ";" & grpField$ & ";ANYCHANGE;A"
End If

Me.vbxCrystal.GroupCondition(1) = grpCond$

'GrpIdx% = comGroup(2).ListIndex
'Select Case GrpIdx%
'   '  Case 0: grpField$ = "{tblPerform.TB_DESC}"
'     Case 1
''        If optWhatDat(0) Then
''            grpField$ = "{HR_SALARY_HISTORY.SH_NEXTDAT}"
''        Else
''            grpField$ = "{HR_PERFORM_HISTORY.PH_PNEXT}"
''        End If
'End Select
'grpCond$ = "GROUP" & CStr(3) & ";" & grpField$ & ";ANYCHANGE;A"
'Me.vbxCrystal.GroupCondition(2) = grpCond$


Cri_Sorts = z% ' next section number to format

End Function

Private Function CriCheck()
Dim x%

CriCheck = False


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
If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
    If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then                               'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpDateRange(0).SetFocus                                         '
        Exit Function                                                       '
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
glbOnTop = Me.name

Screen.MousePointer = HOURGLASS
Call comGrpLoad
Call setRptCaption(Me)

Call INI_Controls(Me)

Me.Caption = lStr(Me.Caption)

lblBCode.Caption = lStr(lblBCode.Caption)


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


