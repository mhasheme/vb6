VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPosDivDeptLnk 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Position, Division and Department Link"
   ClientHeight    =   8340
   ClientLeft      =   1485
   ClientTop       =   885
   ClientWidth     =   10500
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8340
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   9960
      Top             =   6360
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   18
      Top             =   7680
      Width           =   10500
      _Version        =   65536
      _ExtentX        =   18521
      _ExtentY        =   1164
      _StockProps     =   15
      ForeColor       =   0
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
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   60
         TabIndex        =   8
         Tag             =   "Select Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Tag             =   "Close and exit screen"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1635
         TabIndex        =   9
         Tag             =   "Edit the information above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2445
         TabIndex        =   10
         Tag             =   "Save the changes made"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4125
         TabIndex        =   12
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4920
         TabIndex        =   13
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   6000
         TabIndex        =   15
         Tag             =   "Print the Province listing report"
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   3675
      TabIndex        =   4
      Tag             =   "Find specific record"
      Top             =   6930
      Width           =   735
   End
   Begin VB.TextBox txtFindKey 
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
      Left            =   1980
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "00-Search Code"
      Top             =   6960
      Width           =   1500
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "PD_COMPNO"
      DataSource      =   "Data1"
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
      Height          =   285
      Left            =   9960
      MaxLength       =   3
      TabIndex        =   17
      Text            =   "001"
      Top             =   6960
      Visible         =   0   'False
      Width           =   615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "FePosDivDeptLnk.frx":0000
      Height          =   5145
      Left            =   120
      OleObjectBlob   =   "FePosDivDeptLnk.frx":0014
      TabIndex        =   16
      Tag             =   "Province Listings"
      Top             =   120
      Width           =   10215
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10200
      Top             =   7200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      ReportSource    =   1
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "PD_DIV"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Tag             =   "00-Division"
      Top             =   5940
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "PD_DEPTNO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Tag             =   "00-Department"
      Top             =   6360
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpJob 
      DataField       =   "PD_JOB"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Tag             =   "01-Position code"
      Top             =   5520
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   5
   End
   Begin Threed.SSOption optSort 
      Height          =   255
      Index           =   0
      Left            =   1620
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7320
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Position"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSOption optSort 
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   6
      Top             =   7320
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Division"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSOption optSort 
      Height          =   255
      Index           =   2
      Left            =   5340
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7320
      Width           =   2175
      _Version        =   65536
      _ExtentX        =   3836
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Department"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   23
      Top             =   7005
      Width           =   1725
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sort/Search By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   7350
      Width           =   1320
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Position"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   5565
      Width           =   1005
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   5985
      Width           =   1485
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   6405
      Width           =   1485
   End
End
Attribute VB_Name = "frmPosDivDeptLnk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew As Integer
Dim fglbNewRec%, xOldCode As String
Dim fglbID As Integer
Dim xLinkItem As String
Dim xPos, xDiv, xDept As String
'Dim rsGrid As ADODB.Recordset

Private Function chkPos_DivDept()
Dim Prov$, Msg$
Dim rsPosGrpCat As New ADODB.Recordset
Dim SQLQ As String

chkPos_DivDept = False

On Error GoTo chkPos_DivDept_Err

If Len(clpJob) < 1 Then
    MsgBox lblTitle(0).Caption & " is a required field"
    clpJob.SetFocus
    Exit Function
End If
If Len(clpDiv) < 1 Then
    MsgBox lblTitle(1).Caption & " is a required field"
    clpDiv.SetFocus
    Exit Function
End If
If Len(clpDept) < 1 Then
    MsgBox lblTitle(2).Caption & " is a required field"
    clpDept.SetFocus
    Exit Function
End If

'Check if this combination of Position, Division and Department already exists
Set rsPosGrpCat = Nothing
SQLQ = "SELECT * FROM HR_JOB_DIVDEPT_LINK"
SQLQ = SQLQ & " WHERE PD_JOB = '" & clpJob.Text & "'"
SQLQ = SQLQ & " AND PD_DIV = '" & clpDiv.Text & "'"
SQLQ = SQLQ & " AND PD_DEPTNO = '" & clpDept.Text & "'"
rsPosGrpCat.Open SQLQ, gdbAdoIhr001, adOpenStatic
If fglbNewRec% = True Then
    If Not rsPosGrpCat.EOF Then
        'Combination already exist
        MsgBox "This " & lblTitle(0).Caption & "/" & lblTitle(1).Caption & "/" & lblTitle(2).Caption & " link already exists."
        clpJob.SetFocus
        Exit Function
    End If
    rsPosGrpCat.Close
Else
    If (xPos <> clpJob.Text) Or (xDiv <> clpDiv.Text) Or (xDept <> clpDept.Text) Then
        If Not rsPosGrpCat.EOF Then
            'Combination already exist
            MsgBox "This " & lblTitle(0).Caption & ", " & lblTitle(1).Caption & " and " & lblTitle(2).Caption & " link already exists."
            clpJob.SetFocus
            Exit Function
        End If
        rsPosGrpCat.Close
    End If
End If

chkPos_DivDept = True

Exit Function

chkPos_DivDept_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HR_JOB_DIVDEPT_LINK", "chkPos_DivDept")
Resume Next

End Function


Private Sub cmdCancel_Click()
Dim bk
'On Error GoTo Can_Err

Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Data1.Refresh
End If

Call modSTUPD(False)  ' reset screen's attributes

cmdClose.SetFocus

fglbNewRec% = False
xPos = ""
xDiv = ""
xDept = ""

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_JOB_DIVDEPT_LINK", "Cancel")
Resume Next

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Msg As String, a%

On Error GoTo DelErr

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Exit Sub
End If

Data1.Recordset.Delete

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh

'Set rsGrid = Data1.Recordset.Clone
'vbxTrueGrid.FetchRowStyle = True

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Single", "Delete")
Call RollBack '09June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()

Dim SQLQ$

txtFindKey.SetFocus 'added by Marlon Cowan 9/16/97

If Len(txtFindKey) > 0 Then
    If optSort(0) Then
        SQLQ$ = "PD_JOB like '" & txtFindKey.Text & "%'"
    ElseIf optSort(1) Then
        SQLQ$ = "PD_DIV like '" & txtFindKey.Text & "%'"
    ElseIf optSort(2) Then
        SQLQ$ = "PD_DEPTNO like '" & txtFindKey.Text & "%'"
    End If
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ$
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

'If Len(txtFindDesc) > 0 Then
'    SQLQ$ = "DESCR >= '" & txtFindDesc.Text & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ$
'    If Data1.Recordset.EOF Then
'        Data1.Refresh
'    Else
'        txtFindDesc = ""
'    End If
'    Exit Sub
'End If

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
xPos = clpJob.Text
xDiv = clpDiv.Text
xDept = clpDept.Text

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '09June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr

glbCodeRef = True

Data1.Recordset.AddNew

txtComp.Text = glbCompNo
xPos = ""
xDiv = ""
xDept = ""

fglbNewRec% = True

Call modSTUPD(True)

clpJob.SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_JOB_DIVDEPT_LINK", "AddNew")
Resume Next

End Sub

Private Sub cmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim Desc As String
Dim ProvCode

On Error GoTo OK_Err

If Not chkPos_DivDept() Then Exit Sub

Data1.Recordset("PD_COMPNO") = txtComp
Data1.Recordset("PD_LDATE") = Format(Now, "SHORT DATE")
Data1.Recordset("PD_LTIME") = Time$
Data1.Recordset("PD_LUSER") = glbUserID
Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("PD_ID")
    Data1.Refresh
    Data1.Recordset.Find "PD_ID=" & fglbID & " "
    
    'Set rsGrid = Data1.Recordset.Clone
    'vbxTrueGrid.FetchRowStyle = True
End If

fglbNewRec% = False

Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption

Resume Next
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_JOB_DIVDEPT_LINK", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Position, " & lStr("Division") & " and " & lStr("Department") & " Link"
Me.vbxCrystal.WindowTitle = "Position, " & lStr("Division") & " and " & lStr("Department") & " Link Report"
Me.vbxCrystal.BoundReportHeading = "Position, " & lStr("Division") & " and " & lStr("Department") & " Link"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()
glbPos = Data1.Recordset("PD_JOB")
Screen.MousePointer = DEFAULT
Unload Me
End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = ErrorNumber
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_OHS_RLINK_EVENT", "SELECT")
'
'End Sub

Private Sub Form_Load()
Dim SQLQ As String

    Screen.MousePointer = HOURGLASS
    
    Data1.ConnectionString = glbAdoIHRDB
    'SQLQ = "SELECT * FROM HR_JOB_DIVDEPT_LINK "
    'SQLQ = SQLQ & " ORDER BY PD_JOB,PD_DIV,PD_DEPTNO "
    
    SQLQ = "SELECT *, (SELECT JB_DESCR FROM HRJOB WHERE JB_CODE = PD_JOB) AS PD_JOBDESC, "
    SQLQ = SQLQ & " (SELECT DIVISION_NAME FROM HR_DIVISION WHERE DIV = PD_DIV) AS PD_DIVDESC,"
    SQLQ = SQLQ & " (SELECT DF_NAME FROM HRDEPT WHERE DF_NBR = PD_DEPTNO) AS PD_DEPTDESC"
    SQLQ = SQLQ & " FROM HR_JOB_DIVDEPT_LINK"
    SQLQ = SQLQ & " ORDER BY PD_JOB,PD_DIV,PD_DEPTNO "
    
    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    'Set rsGrid = Data1.Recordset.Clone
    'vbxTrueGrid.FetchRowStyle = True
    
    Me.Caption = "Position, " & lStr("Division") & " and " & lStr("Department") & " Link"

    Call setCaption(lblTitle(1))
    Call setCaption(lblTitle(2))
    
    optSort(1).Caption = lblTitle(1).Caption
    optSort(2).Caption = lblTitle(2).Caption
    vbxTrueGrid.Columns(2).Caption = lblTitle(1).Caption
    vbxTrueGrid.Columns(4).Caption = lblTitle(2).Caption
    vbxTrueGrid.Columns(3).Caption = lblTitle(1).Caption & " Name"
    vbxTrueGrid.Columns(5).Caption = lblTitle(2).Caption & " Name"

    Call modSTUPD(False)            'Jaddy 10/18/99
    
    Call INI_Controls(Me)
    
    Screen.MousePointer = DEFAULT
                                
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub modSTUPD(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

cmdOK.Enabled = TF          'May99 js
cmdCancel.Enabled = TF      '
cmdClose.Enabled = FT
cmdPrint.Enabled = FT       '
cmdFind.Enabled = FT        '
cmdSelect.Enabled = FT
cmdDelete.Enabled = FT
If gSec_Upd_Job_Master Then  '
    cmdModify.Enabled = FT      '
    cmdNew.Enabled = FT         '
    cmdDelete.Enabled = FT      '
Else
    cmdModify.Enabled = False      '
    cmdNew.Enabled = False   '
    cmdDelete.Enabled = False
End If

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
clpJob.Enabled = TF
clpDiv.Enabled = TF
clpDept.Enabled = TF
'txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
vbxTrueGrid.Enabled = FT
optSort(0).Enabled = FT
optSort(1).Enabled = FT
optSort(2).Enabled = FT

'If glbDivInhSel Then
'    cmdSelect.Enabled = False
'End If

End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub optSort_Click(Index As Integer, Value As Integer)
    lblTitle(3).Caption = optSort(Index).Caption
    If optSort(0) Then
        Call vbxTrueGrid_HeadClick(0)
    ElseIf optSort(1) Then
        Call vbxTrueGrid_HeadClick(2)
    ElseIf optSort(2) Then
        Call vbxTrueGrid_HeadClick(4)
    End If
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
    'If ColIndex <> 1 And ColIndex <> 3 And ColIndex <> 5 Then
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
                       
        'SQLQ = "SELECT * FROM HR_JOB_DIVDEPT_LINK "
        'SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
        SQLQ = "SELECT *, (SELECT JB_DESCR FROM HRJOB WHERE JB_CODE = PD_JOB) AS PD_JOBDESC, "
        SQLQ = SQLQ & " (SELECT DIVISION_NAME FROM HR_DIVISION WHERE DIV = PD_DIV) AS PD_DIVDESC,"
        SQLQ = SQLQ & " (SELECT DF_NAME FROM HRDEPT WHERE DF_NBR = PD_DEPTNO) AS PD_DEPTDESC"
        SQLQ = SQLQ & " FROM HR_JOB_DIVDEPT_LINK"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
    
        'Set rsGrid = Data1.Recordset.Clone
        'vbxTrueGrid.FetchRowStyle = True
    'End If
End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

'Public Property Let LinkItem(vData As String)
'    xLinkItem = vData
'End Property
'
'Public Property Get LinkItem() As String
'    LinkItem = xLinkItem
'End Property

