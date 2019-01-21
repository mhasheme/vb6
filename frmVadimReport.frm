VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#60.0#0"; "IHRCTRLS.OCX"
Begin VB.Form frmVadimReport 
   Caption         =   "Vadim Intergradation Report"
   ClientHeight    =   6015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   7515
   WindowState     =   2  'Maximized
   Begin VB.Frame frmAT 
      Height          =   435
      Left            =   120
      TabIndex        =   5
      Top             =   540
      Visible         =   0   'False
      Width           =   5115
      Begin VB.OptionButton optAT 
         Caption         =   "Active Employee"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   150
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.OptionButton optAT 
         Caption         =   "Terminated Employee"
         Height          =   255
         Index           =   1
         Left            =   2490
         TabIndex        =   6
         Top             =   150
         Width           =   2175
      End
   End
   Begin VB.ComboBox cmbUpload 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2415
      TabIndex        =   3
      Tag             =   "Choose Upload flag."
      Text            =   "Combo1"
      Top             =   2100
      Width           =   975
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      Height          =   285
      Left            =   2100
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.DateLookup dlpTo 
      Height          =   285
      Left            =   2100
      TabIndex        =   1
      Tag             =   "40-Date upto and including this date forward"
      Top             =   1710
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFrom 
      Height          =   285
      Left            =   2100
      TabIndex        =   2
      Tag             =   "40-Date from and including this date forward"
      Top             =   1380
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7590
      Top             =   6300
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
      Left            =   2430
      TabIndex        =   4
      Tag             =   "Page break after Employee changes"
      Top             =   2940
      Visible         =   0   'False
      Width           =   225
      _Version        =   65536
      _ExtentX        =   397
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Page Break"
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   27.01
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
      Left            =   6990
      Top             =   6300
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
      Left            =   2100
      TabIndex        =   8
      Tag             =   "10-Enter Employee Number"
      Top             =   1050
      Visible         =   0   'False
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6595
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpPP 
      DataField       =   "SH_PAYP"
      Height          =   285
      Left            =   2115
      TabIndex        =   9
      Tag             =   "00-Enter pay period code"
      Top             =   2520
      Visible         =   0   'False
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPP"
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number  "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   17
      Top             =   1080
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   1425
      Width           =   870
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Upload Flag"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   210
      TabIndex        =   15
      Top             =   2130
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Page Break on Employee"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   210
      TabIndex        =   14
      Top             =   2940
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1560
      TabIndex        =   13
      Top             =   1425
      Width           =   420
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1590
      TabIndex        =   12
      Top             =   1740
      Width           =   240
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Facility"
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Period"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   10
      Top             =   2520
      Visible         =   0   'False
      Width           =   930
   End
End
Attribute VB_Name = "frmVadimReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DeletedRecs As Long

Private Function chkVadim()
Dim dd&
Dim Msg$, DgDef As Variant, Response%
chkVadim = False
On Error GoTo chkEOTHERE_Err


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
  dd& = DateDiff("d", CVDate(dlpFrom.Text), CVDate(dlpTo.Text))
  If dd& < 0 Then
    MsgBox "From date must be earlier than To Date"
    dlpFrom.SetFocus
    Exit Function
  End If
End If
'If Not elpEEID.ListChecker Then
'    Exit Function
'End If

chkVadim = True
Exit Function

chkEOTHERE_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkVadim", "Vadim", "Update")
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



Private Sub Cri_PP()
    Dim PPCri As String
    
    If Len(clpPP.Text) > 0 Then
      PPCri = "{HR_SALARY_HISTORY.SH_PAYP} in ['" & clpPP.Text & "'] "
      PPCri = PPCri & "AND {HR_SALARY_HISTORY.SH_CURRENT} "
      If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
      glbstrSelCri = glbstrSelCri & PPCri
    End If
End Sub

Public Sub cmdDelete_Click()
Dim a As Integer
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%
Title$ = "Mass Vadim File Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

Screen.MousePointer = HOURGLASS
gdbPayroll.Execute "DELETE FROM SY_INTERFACE"
gdbPayroll.Execute "DELETE FROM SY_INYTERFACE_BATCH"
Screen.MousePointer = DEFAULT

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Other Earnings", "Delete")
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub



Public Sub cmdPrint_Click()
Dim x%

On Error GoTo PrntErr
Screen.MousePointer = HOURGLASS
If chkVadim() Then
  If Not PrtForm("Vadim Master Update Criteria", Me) Then
    Exit Sub
  End If
  x% = Cri_SetAll()
  Me.vbxCrystal.Destination = 1
  MDIMain.Timer1.Enabled = False
  Me.vbxCrystal.Action = 1
  vbxCrystal.Reset
  MDIMain.Timer1.Enabled = True
End If
Screen.MousePointer = DEFAULT
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
Dim x%

On Error GoTo ViewErr

Screen.MousePointer = HOURGLASS

If chkVadim() Then
  x% = Cri_SetAll()
  Me.vbxCrystal.Destination = 0
  MDIMain.Timer1.Enabled = False
  Me.vbxCrystal.Action = 1
  vbxCrystal.Reset
  MDIMain.Timer1.Enabled = True
End If
Screen.MousePointer = DEFAULT
Exit Sub

ViewErr:
MsgBox "CRW ERROR : " & Chr(10) & "[" & Str(Err) & "] : " & Me.vbxCrystal.LastErrorString

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdView_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub



Private Sub Cri_EE()
Dim EECri As String

If Len(elpEEID.Text) > 0 Then
  EECri = "{HRAUDIT.AU_EMPNBR} in [" & getEmpnbr(elpEEID.Text) & "] "
  If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
  glbstrSelCri = glbstrSelCri & EECri
End If

End Sub

Private Sub Cri_FTDates()
Dim TempCri As String
Dim dtYYY%, dtMM%, dtDD%
Dim x%

If Len(dlpFrom.Text) = 0 And Len(dlpTo.Text) = 0 Then Exit Sub
TempCri = "({SY_INYTERFACE_BATCH.PROCESS_DATE} "
If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
  dtYYY% = Year(dlpFrom.Text)
  dtMM% = Month(dlpFrom.Text)
  dtDD% = Day(dlpFrom.Text)
  TempCri = TempCri & " in Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ") "
  dtYYY% = Year(dlpTo.Text)
  dtMM% = Month(dlpTo.Text)
  dtDD% = Day(dlpTo.Text)
  TempCri = TempCri & " to Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
Else
  If Len(dlpFrom.Text) > 0 Then
    TempCri = TempCri & " >= "
    dtYYY% = Year(dlpFrom.Text)
    dtMM% = Month(dlpFrom.Text)
    dtDD% = Day(dlpFrom.Text)
    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
  End If
  If Len(dlpTo.Text) > 0 Then
    TempCri = TempCri & " <= "
    dtYYY% = Year(dlpTo.Text)
    dtMM% = Month(dlpTo.Text)
    dtDD% = Day(dlpTo.Text)
    TempCri = TempCri & " Date(" & dtYYY% & ", " & dtMM% & ", " & dtDD% & ")) "
  End If
End If
If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
glbstrSelCri = glbstrSelCri & TempCri

End Sub

Private Function Cri_SetAll()
Dim x%

Cri_SetAll = False

On Error GoTo modSetCriteria_Err
Screen.MousePointer = HOURGLASS

glbiOneWhere = False
glbstrSelCri = ""
'Call glbCri_DeptUN("")
'Call Cri_Div
'Call Cri_EE
'Call Cri_PP
Call Cri_FTDates
Call Cri_Upload
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGVADIM.rpt"

Me.vbxCrystal.SelectionFormula = glbstrSelCri
Me.vbxCrystal.Connect = "ODBC;DSN=IHRVADIM;UID=" & InterUserName & ";PWD=" & InterUserPassword

'If chkPage Then
'  Me.vbxCrystal.SectionFormat(0) = "GH1;T;F;X;X;X;X;X"
'  Me.vbxCrystal.SectionFormat(1) = "GF1;X;X;T;X;X;X;X"
'  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = 'Employee:'"
'  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = {@EFullName}"
'  Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = 'Number:'"
'  'Me.vbxCrystal.Formulas(3) = "DESCGROUP4 = {HRAUDIT.AU_EMPNBR}"
'Else
'  Me.vbxCrystal.SectionFormat(0) = "GH1;T;X;X;X;X;X;X"
'  Me.vbxCrystal.SectionFormat(1) = "GF1;X;F;X;X;X;X;X"
'  Me.vbxCrystal.Formulas(0) = "DESCGROUP1 = ''"
'  Me.vbxCrystal.Formulas(1) = "DESCGROUP2 = ''"
'  Me.vbxCrystal.Formulas(2) = "DESCGROUP3 = ''"
'  Me.vbxCrystal.Formulas(3) = "lblEMPNO = ''"
'End If
Me.vbxCrystal.WindowTitle = "Vadim Master File Report"
Cri_SetAll = True
Screen.MousePointer = DEFAULT
Exit Function

modSetCriteria_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "VAdim Master", "Vadim Report", "Select")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Sub Cri_Upload()
Dim EECri As String
If cmbUpload.ListIndex > 0 Then
  If cmbUpload.ListIndex = 1 Then
    EECri = "{SY_INYTERFACE_BATCH.PROCESS_CODE} = 'Y' "
  End If
  If cmbUpload.ListIndex = 2 Then
    EECri = "{SY_INYTERFACE_BATCH.PROCESS_CODE} = 'N' "
  End If
  If Len(glbstrSelCri) > 3 Then glbstrSelCri = glbstrSelCri & " AND "
  glbstrSelCri = glbstrSelCri & EECri
End If
End Sub



Private Sub Form_Activate()
Call SET_UP_MODE

End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim SQLQ As String
Screen.MousePointer = HOURGLASS

cmbUpload.AddItem "All"
cmbUpload.AddItem "Yes"
cmbUpload.AddItem "No"
cmbUpload.ListIndex = 0

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
    Set frmVadimReport = Nothing  'carmen may 2000
End Sub



Private Sub Cri_Div()

Dim DivCri As String
Dim countr   As Integer

If Len(clpDIV.Text) > 0 Then
    DivCri = "(RIGHT(TOTEXT({HRAUDIT.AU_EMPNBR},0),3) = '" & clpDIV.Text & "')"
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
    Deleteble = App.Path = "C:\SSWORK\IHR73"
End Property

Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
MDIMain.MainToolBar.ButtonS(10).Visible = True
MDIMain.MainToolBar.ButtonS(10).Enabled = True

End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub



