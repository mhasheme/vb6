VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmAttAUDIT 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attendance Audit Report"
   ClientHeight    =   7350
   ClientLeft      =   4380
   ClientTop       =   3915
   ClientWidth     =   11115
   DrawMode        =   1  'Blackness
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
   Icon            =   "fxAttaudit.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7350
   ScaleWidth      =   11115
   Tag             =   "Attendance Audit Update"
   WindowState     =   2  'Maximized
   Begin VB.Frame frmAT 
      Height          =   435
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   5115
      Begin VB.OptionButton optAttend 
         Caption         =   "Attendance"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   150
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optAttend 
         Caption         =   "Attendance History"
         Height          =   255
         Index           =   1
         Left            =   2490
         TabIndex        =   1
         Top             =   150
         Width           =   2175
      End
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "Final sorting of records - no totals"
      Top             =   5760
      Width           =   2325
   End
   Begin INFOHR_Controls.DateLookup dlpTo 
      Height          =   285
      Left            =   2190
      TabIndex        =   4
      Tag             =   "40-Date upto and including this date forward"
      Top             =   1650
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFrom 
      Height          =   285
      Left            =   2190
      TabIndex        =   3
      Tag             =   "40-Date from and including this date forward"
      Top             =   1320
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7680
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSCheck chkPage 
      Height          =   225
      Left            =   2520
      TabIndex        =   6
      Tag             =   "Page break after Employee changes"
      Top             =   5280
      Width           =   225
      _Version        =   65536
      _ExtentX        =   397
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Page Break"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
      Font3D          =   3
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   6120
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
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2190
      TabIndex        =   2
      Tag             =   "10-Enter Employee Number"
      Top             =   990
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6595
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.EmployeeLookup elpUser 
      Height          =   315
      Left            =   2190
      TabIndex        =   5
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      ShowDescription =   0   'False
      RefreshDescriptionWhen=   2
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "User"
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
      Left            =   300
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
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
      Left            =   300
      TabIndex        =   13
      Top             =   5790
      Width           =   660
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   1680
      Width           =   240
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   1650
      TabIndex        =   11
      Top             =   1365
      Width           =   420
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break on Employee"
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
      Left            =   300
      TabIndex        =   10
      Top             =   5280
      Width           =   1800
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
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
      Left            =   300
      TabIndex        =   9
      Top             =   1365
      Width           =   870
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number  "
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
      Left            =   300
      TabIndex        =   8
      Top             =   1020
      Width           =   1380
   End
End
Attribute VB_Name = "frmAttAUDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DeletedRecs As Long

Private Function chkAudit()
Dim dd As Long

chkAudit = False
On Error GoTo chkEOTHERE_Err
'If glbLinamar Then
'    If Len(clpDiv) > 0 Then
'        If clpDiv.Caption = "Unassigned" Then
'            MsgBox "If Facility Entered - they must exist"
'            clpDiv.SetFocus
'            Exit Function
'        End If
'    End If
'Else
'    If Not clpDiv1.ListChecker Then
'        Exit Function
'    End If
'End If

If Len(dlpFrom.Text) > 0 Then
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "Invalid From date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If
If Len(dlpTo.Text) > 0 Then
    If Not IsDate(dlpTo.Text) Then
        MsgBox "Invalid To date"
        dlpTo.SetFocus
        Exit Function
    End If
End If
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dd = DateDiff("d", CVDate(dlpFrom.Text), CVDate(dlpTo.Text))
    If dd < 0 Then
        MsgBox "From date must be earlier than To Date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If
If Not elpEEID.ListChecker Then
    Exit Function
End If

chkAudit = True
Exit Function

chkEOTHERE_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkAudit", "ATT_AUDIT", "Update")
Resume Next

End Function

Private Sub chkPage_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbUpload_GotFocus()
    Call SetPanHelp(ActiveControl)
    MDIMain.panHelp(2).Caption = "Req."
End Sub

Public Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub Cri_PP()
'    Dim PPCri As String
'
'    If Len(clpPP.Text) > 0 Then
'      PPCri = "{HR_SALARY_HISTORY.SH_PAYP} in ['" & clpPP.Text & "'] "
'      If glbOracle Then
'        PPCri = PPCri & "AND {HR_SALARY_HISTORY.SH_CURRENT}<>0 "
'      Else
'        PPCri = PPCri & "AND {HR_SALARY_HISTORY.SH_CURRENT} "
'      End If
'      If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
'      glbstrSelCri = glbstrSelCri & PPCri
'    End If
'End Sub

'Public Sub cmdDelete_Click()
'Dim X As Integer
'Dim DgDef, Title As String, Msg As String, Response As Integer
'If glbLinamar Then
'    If Len(clpDiv) = 0 Then
'        MsgBox "Facility is a required field"
'        clpDiv.SetFocus
'        Exit Sub
'    End If
'End If
'Title = "Mass Audit File Delete"
'DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
'Msg = "Are You Sure You Want To Delete ALL records for this criteria?"
'Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
'
'If Response = IDNO Then    ' Evaluate response
'    Exit Sub
'End If
'
'Screen.MousePointer = HOURGLASS
'X = modDelRecs()
'Screen.MousePointer = DEFAULT
'If DeletedRecs = 0 Then
'    MsgBox "No records found for given selection criteria."
'Else
'    MsgBox DeletedRecs & " records deleted successfully"
'End If
'
'Exit Sub
'
'Del_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Other Earnings", "Delete")
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
On Error GoTo PrntErr
Dim X As Integer

Screen.MousePointer = HOURGLASS
If chkAudit() Then
    If Not PrtForm("Attendance Audit Report Criteria", Me) Then
        Exit Sub
    End If
    ' cmdView.Enabled = False
    ' cmdPrint.Enabled = False
    ' cmdDelete.Enabled = False
     X = Cri_SetAll()
     Me.vbxCrystal.Destination = 1
     MDIMain.Timer1.Enabled = False
     Me.vbxCrystal.Action = 1
     vbxCrystal.Reset
     MDIMain.Timer1.Enabled = True
    '  cmdView.Enabled = True
    '  cmdPrint.Enabled = True
    '  If gSec_Upd_Audit Then cmdDelete.Enabled = True
End If

Screen.MousePointer = DEFAULT

Exit Sub

PrntErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdView_Click()
Dim X As Integer

On Error GoTo ViewErr

Screen.MousePointer = HOURGLASS

If chkAudit() Then
    '  cmdView.Enabled = False
    '  cmdPrint.Enabled = False
    '  cmdDelete.Enabled = False
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    X = Cri_SetAll()
    Me.vbxCrystal.Destination = 0
    MDIMain.Timer1.Enabled = False
    Me.vbxCrystal.Action = 1
    vbxCrystal.Reset
    MDIMain.Timer1.Enabled = True
    '  cmdView.Enabled = True
    '  cmdPrint.Enabled = True
    '  If gSec_Upd_Audit Then cmdDelete.Enabled = True
End If

Screen.MousePointer = DEFAULT

Exit Sub

ViewErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & str(Err) & "] : " & Me.vbxCrystal.LastErrorString

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
    If optAttend(0) <> 0 Then
        EECri = "{HR_ATTENDANCE.AD_EMPNBR} in [" & getEmpnbr(elpEEID.Text) & "] "
    Else
        EECri = "{HR_ATTENDANCE_HISTORY.AH_EMPNBR} in [" & getEmpnbr(elpEEID.Text) & "] "
    End If
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    glbstrSelCri = glbstrSelCri & EECri
End If

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY As Integer, dtMM As Integer, dtDD As Integer


If Len(dlpFrom.Text) = 0 And Len(dlpTo.Text) = 0 Then Exit Sub
If optAttend(0) <> 0 Then
    TempCri = "({HR_ATTENDANCE.AD_LDATE} "
Else
    TempCri = "({HR_ATTENDANCE_HISTORY.AH_LDATE} "
End If
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dtYYY = Year(dlpFrom.Text)
    dtMM = month(dlpFrom.Text)
    dtDD = Day(dlpFrom.Text)
    TempCri = TempCri & " in Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ") "
    dtYYY = Year(dlpTo.Text)
    dtMM = month(dlpTo.Text)
    dtDD = Day(dlpTo.Text)
    TempCri = TempCri & " to Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
Else
    If Len(dlpFrom.Text) > 0 Then
        TempCri = TempCri & " >= "
        dtYYY = Year(dlpFrom.Text)
        dtMM = month(dlpFrom.Text)
        dtDD = Day(dlpFrom.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
    If Len(dlpTo.Text) > 0 Then
        TempCri = TempCri & " <= "
        dtYYY = Year(dlpTo.Text)
        dtMM = month(dlpTo.Text)
        dtDD = Day(dlpTo.Text)
        TempCri = TempCri & " Date(" & dtYYY & ", " & dtMM & ", " & dtDD & ")) "
    End If
End If
If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
glbstrSelCri = glbstrSelCri & TempCri

End Sub

Private Function Cri_SetAll()
On Error GoTo modSetCriteria_Err
Dim X As Integer

Cri_SetAll = False

Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
Call glbCri_DeptUN("")
' call cri models set both glbiONeWhere and strSelCri
'If glbLinamar Then
'    Call Cri_Div
'Else
'    Call Cri_Div1
'End If
Call Cri_EE
'Call Cri_PP
Call Cri_FTDates
'Call Cri_Upload
'Call Cri_Checks
Call Cri_Sorts
Call Cri_User
Call setRptLabel(Me, 2)
If optAttend(0) <> 0 Then
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAttAudit.rpt"
Else
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rzAttHisAudit.rpt"
End If

Me.vbxCrystal.SelectionFormula = glbstrSelCri
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "PWD=petman;"
'    If optAT(0) <> 0 Then
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
'    Else
'        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
'    End If
    Me.vbxCrystal.DataFiles(1) = glbIHRDB
    Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
    Me.vbxCrystal.DataFiles(3) = glbIHRDB
End If
If chkPage Then
  Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
  Me.vbxCrystal.SectionFormat(1) = "GF1;X;X;T;X;X;X;X"
Else
  Me.vbxCrystal.SectionFormat(0) = "GH1;T;X;X;X;X;X;X"
  Me.vbxCrystal.SectionFormat(1) = "GF1;X;F;X;X;X;X;X"
End If

'If glbPayWeb Then
'  Me.vbxCrystal.Formulas(10) = "xWCB = IF LENGTH ({HRAUDIT.AU_WCB}) > 0 THEN  {HRAUDIT.AU_TYPE} + '                          E.I. Reduced Rate'"
'End If
If glbWFC Then 'Ticket #12867
    Me.vbxCrystal.Formulas(10) = "WFCNoEXECuser = " & glbNoEXEC & " "
End If

' window title if appropriate
If optAttend(0) <> 0 Then
    Me.vbxCrystal.WindowTitle = "Attendance Audit Report"
Else
    Me.vbxCrystal.WindowTitle = "Attendance History Audit Report"
End If

Cri_SetAll = True
'For x = 0 To 1000
'    If Me.vbxCrystal.Formulas(x) <> "" Then Debug.Print Me.vbxCrystal.Formulas(x)
'Next
Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Attendance Audit", "ATT_AUDIT Report", "Select")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

'Private Sub Cri_Upload()
'Dim EECri As String
'If cmbUpload.ListIndex > 0 Then
'  If cmbUpload.ListIndex = 1 Then
'    EECri = "{HRAUDIT.AU_UPLOAD} = 'Y' "
'  End If
'  If cmbUpload.ListIndex = 2 Then
'    EECri = "{HRAUDIT.AU_UPLOAD} = 'N' "
'  End If
'  If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
'  glbstrSelCri = glbstrSelCri & EECri
'End If
'End Sub


Private Sub Form_Activate()
glbOnTop = "FRMATTAUDIT"
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim SQLQ As String

glbOnTop = "FRMATTAUDIT"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'If glbLinamar Then
'    lblDiv.Visible = False
'    clpDiv1.Visible = False
'Else
'    Call setCaption(lblDiv)
'End If

Screen.MousePointer = HOURGLASS
'If glbLinamar Then
'    clpPP.Visible = False
'    lblPP.Visible = False
'End If
Data1.ConnectionString = glbAdoIHRAUDIT

'Hemu - Talked to Jerry about this read - not really required. It's taking time to load
'the page - 09/12/2008
'SQLQ = "SELECT AU_EMPNBR FROM HRAUDIT WHERE AU_EMPNBR IN(SELECT ED_EMPNBR FROM HREMP "
'SQLQ = SQLQ & in_SQL(glbIHRDB)
'SQLQ = SQLQ & " WHERE " & glbSeleDeptUn & ")"
'Data1.RecordSource = SQLQ
'Data1.Refresh

'If Data1.Recordset.EOF And Data1.Recordset.EOF Then
'  MsgBox "ACTIVE AUDIT FILE IS EMPTY"
'  Screen.MousePointer = DEFAULT
'End If

'cmbUpload.AddItem "All"
'cmbUpload.AddItem "Yes"
'cmbUpload.AddItem "No"
'cmbUpload.ListIndex = 0

comGroup.Clear
comGroup.AddItem "Date Changed"
comGroup.AddItem "Employee Number"
comGroup.AddItem "Employee Name"
comGroup.ListIndex = 0

If Not gSec_Upd_Audit Then     'May99 js
'    cmdDelete.Enabled = False   '
End If                          '
'If glbLinamar Then
'    lblTitle(0).Visible = True
'    clpDiv.Visible = True
'    frmAT.Visible = True
'End If
elpUser.LookupType = 2

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
    Set frmAttAUDIT = Nothing  'carmen may 2000
End Sub

'Private Function modDelRecs()
''''On Error GoTo cmdDel_Err
'Dim SQLQ As String, SQLW As String, SQL1 As String, SQLQ1 As String
'Dim TmpDeletedRecs As Long, DeletedRecs1 As Long, TmpDeletedRecs1 As Long, TmpDeletedRecs2 As Long, DeletedEmp0Recs As Long
'Dim SQLQ2, SQLQ_0 As String
'
'modDelRecs = False
'
'glbstrSelCri = ""
'Screen.MousePointer = HOURGLASS
'
'SQLQ = "Delete FROM HRAUDIT WHERE 1=1 "
'
'' do selection for pay period if they entered one
'If Len(clpPP.Text) > 0 Then
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM HR_SALARY_HISTORY "
'    If Not glbSQL Then
'        SQLQ = SQLQ & in_SQL(glbIHRDB)
'    End If
'    SQLQ = SQLQ & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
'End If
'
'' pay period selection end
'If glbLinamar Then
'    ' do selection for only emps we have security for
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
'    SQLQ = SQLQ & in_SQL(glbIHRDB)
'    SQLW = "WHERE " & glbSeleDeptUn & ")"
'Else
'    SQLW = ""
'End If
'
'SQLQ1 = SQLQ
'
'If Len(elpEEID.Text) > 0 Then SQLW = SQLW & " AND AU_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
'If Len(dlpFrom.Text) > 0 Then SQLW = SQLW & " AND AU_LDATE >= " & Date_SQL(dlpFrom.Text)
'If Len(dlpTo.Text) > 0 Then SQLW = SQLW & " AND AU_LDATE <= " & Date_SQL(dlpTo.Text)
'If glbLinamar Then
'    If Len(clpDiv) > 0 Then SQLW = SQLW & " AND RIGHT(AU_EMPNBR,3)=" & clpDiv
'Else
'    If Len(clpDiv1.Text) > 0 Then SQLW = SQLW & " AND AU_DIVUPL IN ('" & getCodes(clpDiv1.Text) & "') "
'End If
'If Len(elpUser.Text) > 0 Then SQLW = SQLW & "AND AU_LUSER = '" & elpUser.Text & "' "
'If cmbUpload.ListIndex > 0 Then
'  If cmbUpload.ListIndex = 1 Then SQLW = SQLW + " AND AU_UPLOAD = 'Y' "
'  If cmbUpload.ListIndex = 2 Then SQLW = SQLW + " AND AU_UPLOAD = 'N' "
'End If
'
'glbstrSelCri = ""
'If glbSQL Or glbOracle Then
'    Call glbCri_DeptUN("")
'    glbstrSelCri = Trim(Replace(Replace(glbstrSelCri, "{", ""), "}", ""))
'    If LCase(Left(Trim(glbstrSelCri), 3)) = "and" Then
'        glbstrSelCri = Mid(glbstrSelCri, 4, Len(glbstrSelCri) - 3)
'    End If
'    glbstrSelCri = " AND (AU_EMPNBR in (SELECT ED_EMPNBR FROM HREMP WHERE " & glbstrSelCri & ") OR AU_EMPNBR in (SELECT ED_EMPNBR FROM Term_HREMP WHERE " & Replace(glbstrSelCri, "HREMP.", "Term_HREMP.") & ")  )"
'
'    SQLW = SQLW & glbstrSelCri
'End If
'
'If glbLinamar Then
'    SQLW = SQLW & " AND AU_TYPE<>'R'"
'End If
'SQLQ = SQLQ & SQLW
'gdbAdoIhr001X.Execute SQLQ, DeletedRecs
'
''--------------------------------------------------------------------------------------------
''Delete Audit records with AU_DIVUPL = blank or null
'If Not glbLinamar Or Len(clpDiv1.Text) > 0 Then
'    SQL1 = ""
'    If Len(elpEEID.Text) > 0 Then SQL1 = SQL1 & " AND AU_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
'    If Len(dlpFrom.Text) > 0 Then SQL1 = SQL1 & " AND AU_LDATE >= " & Date_SQL(dlpFrom.Text)
'    If Len(dlpTo.Text) > 0 Then SQL1 = SQL1 & " AND AU_LDATE <= " & Date_SQL(dlpTo.Text)
'
'    'If Len(clpDiv1.Text) > 0 Then SQL1 = SQL1 & " AND AU_DIVUPL IN ('" & getCodes(clpDiv1.Text) & "') "
'    If Len(clpDiv1.Text) > 0 Then SQL1 = SQL1 & " AND ((AU_DIVUPL IS NULL OR AU_DIVUPL = '') AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_DIV IN ('" & getCodes(clpDiv1.Text) & "')))"
'
'    If Len(elpUser.Text) > 0 Then SQL1 = SQL1 & " AND AU_LUSER = '" & elpUser.Text & "' "
'    If cmbUpload.ListIndex > 0 Then
'      If cmbUpload.ListIndex = 1 Then SQL1 = SQL1 + " AND AU_UPLOAD = 'Y' "
'      If cmbUpload.ListIndex = 2 Then SQL1 = SQL1 + " AND AU_UPLOAD = 'N' "
'    End If
'    SQL1 = SQL1 & glbstrSelCri
'    SQLQ1 = SQLQ1 & SQL1
'    gdbAdoIhr001X.Execute SQLQ1, DeletedRecs1
'End If
''--------------------------------------------------------------------------------------------
'
'' dkostka - 08/20/2001 - Added code to remove records for terminated emps too
'SQLQ = "DELETE FROM HRAUDIT WHERE 1=1 "
'If glbLinamar Then
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT ED_EMPNBR FROM Term_HREMP "
'End If
'SQLQ = SQLQ & SQLW
'
'' do selection for pay period if they entered one
'If Len(clpPP.Text) > 0 Then
'    SQLQ = SQLQ & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM Term_SALARY_HISTORY "
'    SQLQ = SQLQ & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
'End If
'
'SQLQ = SQLQ & glbstrSelCri
'
'' pay period selection end
'gdbAdoIhr001X.Execute SQLQ, TmpDeletedRecs
''DeletedRecs = DeletedRecs + TmpDeletedRecs
'
''--------------------------------------------------------------------------------------------
''Delete Audit records with AU_DIVUPL = blank or null - Terminated employees
'If Not glbLinamar Or Len(clpDiv1.Text) > 0 Then
'    SQLQ2 = "DELETE FROM HRAUDIT WHERE 1=1 "
'    SQLQ2 = SQLQ2 & SQL1
'
'    ' do selection for pay period if they entered one
'    If Len(clpPP.Text) > 0 Then
'        SQLQ2 = SQLQ2 & "AND AU_EMPNBR IN (SELECT SH_EMPNBR FROM Term_SALARY_HISTORY "
'        SQLQ2 = SQLQ2 & "WHERE SH_CURRENT<>0 AND SH_PAYP='" & clpPP.Text & "') "
'    End If
'
'    SQLQ2 = SQLQ2 & glbstrSelCri
'
'    ' pay period selection end
'    gdbAdoIhr001X.Execute SQLQ2, TmpDeletedRecs1
'    'DeletedRecs = DeletedRecs + TmpDeletedRecs1 + DeletedRecs1
'End If
''--------------------------------------------------------------------------------------------
'
''Ticket #15576 v7.8 make HRAUDIT2 for customers
''If glbCompSerial = "S/N - 2380W" Then   'VitalAire Ticket #12142
''HRAUDIT2
'    TmpDeletedRecs = 0
'    SQLQ = Replace(SQLQ, "HRAUDIT", "HRAUDIT2")
'    gdbAdoIhr001X.Execute SQLQ, TmpDeletedRecs2
'    'DeletedRecs = DeletedRecs + TmpDeletedRecs + TmpDeletedRecs2
''End If
'
''Ticket #16768
'SQLQ_0 = "DELETE FROM HRAUDIT WHERE AU_EMPNBR = 0"
'gdbAdoIhr001X.Execute SQLQ_0, DeletedEmp0Recs
'
'DeletedRecs = DeletedRecs + DeletedRecs1 + TmpDeletedRecs + TmpDeletedRecs1 + TmpDeletedRecs2
'
'
'modDelRecs = True
'
'Exit Function
'
'cmdDel_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "ATT_AUDIT", "Delete")
'
'Screen.MousePointer = DEFAULT
'
'If gintRollBack% = False Then
'    RollBack
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Function

'Private Sub Cri_Div()
'
'Dim DivCri As String
'
'If Len(clpDiv.Text) > 0 Then
'    DivCri = "(RIGHT(TOTEXT({HR_ATTENDANCE.AD_EMPNBR},0),3) = '" & clpDiv.Text & "')"
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

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = Reports
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Audit
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
'MDIMain.MainToolBar.ButtonS(10).Visible = True
'MDIMain.MainToolBar.ButtonS(10).Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Cri_Checks()
'Added by Bryan 6/Jul/05 for ticket#8857
    Dim TempCri As String
        
    'If Not glbLinamar Then
    '    If Not clpDiv1.ListChecker Then
    '        Exit Sub
    '    End If
    'End If
    
    'If chkNewHires.Value = vbChecked Then
    '    TempCri = "({HRAUDIT.AU_NEWEMP} = 'Y') "
    'End If
    
    'If chkSalary.Value = vbChecked Then
    '    If Len(TempCri) >= 1 Then
    '        TempCri = TempCri & "AND (isnull({HRAUDIT.AU_SALARY})=false ) "
    '    Else
    '        TempCri = "(isnull({HRAUDIT.AU_SALARY})=false ) "
    '    End If
    'End If
    
    'If chkBenefits.Value = vbChecked Then
    '    If Len(TempCri) >= 1 Then
    '        TempCri = TempCri & "AND (isnull({HRAUDIT.AU_BCODE})=false ) "
    '    Else
    '        TempCri = "(isnull({HRAUDIT.AU_BCODE})=false ) "
    '    End If
    'End If
    
    'If chkNameAdd.Value = vbChecked Then
    '    If Len(TempCri) >= 1 Then
    '        TempCri = TempCri & "AND ((isnull({HRAUDIT.AU_FNAME})=false) OR (isnull({HRAUDIT.AU_SURNAME})=false) OR (isnull({HRAUDIT.AU_ADDR1})=false) OR (isnull({HRAUDIT.AU_ADDR2})=false) OR (isnull({HRAUDIT.AU_CITY})=false) OR (isnull({HRAUDIT.AU_PCODE})=false) OR (isnull({HRAUDIT.AU_PROV})=false) ) "
    '    Else
    '        TempCri = "((isnull({HRAUDIT.AU_FNAME})=false) OR (isnull({HRAUDIT.AU_SURNAME})=false) OR (isnull({HRAUDIT.AU_ADDR1})=false) OR (isnull({HRAUDIT.AU_ADDR2})=false) OR (isnull({HRAUDIT.AU_CITY})=false) OR (isnull({HRAUDIT.AU_PCODE})=false) OR (isnull({HRAUDIT.AU_PROV})=false) ) "
    '    End If
    'End If
    
  If Len(glbstrSelCri) > 3 And Len(TempCri) >= 1 Then glbstrSelCri = glbstrSelCri & " AND "
  glbstrSelCri = glbstrSelCri & TempCri

    
End Sub

Private Sub Cri_Sorts()
'Added by Bryan on Sep 7, 2005 Ticket#9279
    Dim grpField As String
    Dim grpCond As String
    'If optAT(0) <> 0 Then
        Select Case comGroup.ListIndex
            Case 0:
                If optAttend(0) <> 0 Then
                    grpField = "{HR_ATTENDANCE.AD_LDATE}"
                Else
                    grpField = "{HR_ATTENDANCE_HISTORY.AH_LDATE}"
                End If
                If optAttend(0) <> 0 Then
                    grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE.AD_LDATE};ANYCHANGE;A"
                Else
                    grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE_HISTORY.AH_LDATE};ANYCHANGE;A"
                End If
                
                Me.vbxCrystal.GroupCondition(0) = grpCond
                grpCond = "GROUP" & CStr(2) & ";{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(1) = grpCond
                Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Date of Change:'"
                If optAttend(0) <> 0 Then
                    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {HR_ATTENDANCE.AD_LDATE}"
                Else
                    Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {HR_ATTENDANCE_HISTORY.AH_LDATE}"
                End If
                
                Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = ''"
                Me.vbxCrystal.Formulas(3) = "lblEMPNO = ''"
            Case 1:
                If optAttend(0) <> 0 Then
                    grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE.AD_EMPNBR};ANYCHANGE;A"
                Else
                    grpCond = "GROUP" & CStr(1) & ";{HR_ATTENDANCE_HISTORY.AH_EMPNBR};ANYCHANGE;A"
                End If
                
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HRAUDIT.AU_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case 2:
                grpCond = "GROUP" & CStr(1) & ";{@EFullName};ANYCHANGE;A"
                Me.vbxCrystal.GroupCondition(0) = grpCond
'                grpCond = "GROUP" & CStr(2) & ";{HRAUDIT.AU_LDATE};ANYCHANGE;A"
'                Me.vbxCrystal.GroupCondition(1) = grpCond
                grpField = "{@EFullName}"
                Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
                Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
                Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
            Case Else: grpField = "(none)"
        End Select
    'End If

End Sub

Private Sub Cri_User()
Dim EECri As String

If Len(elpUser.Text) > 0 Then
    If optAttend(0) <> 0 Then
        EECri = "LowerCase({HR_ATTENDANCE.AD_LUSER}) ='" & LCase(elpUser.Text) & "' "
    Else
        EECri = "LowerCase({HR_ATTENDANCE_HISTORY.AH_LUSER}) ='" & LCase(elpUser.Text) & "' "
    End If
    If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
    glbstrSelCri = glbstrSelCri & EECri
End If

End Sub

'Private Sub Cri_Div1()
'
'Dim DivCri As String
'Dim countr   As Integer  ' EEList_Snap is definded at form level
'
'
'If Len(clpDiv1.Text) > 0 Then
'    'Hemu 06/02/2004 Begin
'    'DivCri = "({HREMP.ED_DIV} = '" & clpDiv.Text & "')"
'    'If glbOracle Then
'    '    DivCri = "({HREMP.ED_DIV} IN ['" & getCodes(clpDiv1.Text) & "'])"
'    'Else
'    '    DivCri = "({HRAUDIT.AU_DIV} IN ('" & getCodes(clpDiv1.Text) & "'))"
'    'End If
'    'Hemu 06/02/2004 End
'
'    'Ticket #12843
'    'DivCri = "({HRAUDIT.AU_DIVUPL} IN ('" & getCodes(clpDiv1.Text) & "'))"
'    'Ticket #13540 Frank, come AU_DIVUPL values were null or blank, but still showup on the report
'    'DivCri = "(Length({HRAUDIT.AU_DIVUPL})>0  AND ({HRAUDIT.AU_DIVUPL} IN ('" & getCodes(clpDiv1.Text) & "')))"
'    DivCri = "(Length({HRAUDIT.AU_DIVUPL})>0  AND ({HRAUDIT.AU_DIVUPL} IN ['" & getCodes(clpDiv1.Text) & "']))"
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

Private Sub optAT_Click(Index As Integer)
    'Ticket #15483
    If Index = 1 Then
        elpEEID.LookupType = TERM
    Else
        elpEEID.LookupType = 0  '0 = ACTIVE. I cannot put as ACTIVE because it's changing to "Active" and that does not switch the lookup to ACTIVE employees
    End If
End Sub
