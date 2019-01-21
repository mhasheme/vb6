VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSEmpTypeMatrix 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form 7 Employee Type Matrix"
   ClientHeight    =   5955
   ClientLeft      =   1485
   ClientTop       =   885
   ClientWidth     =   8250
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
   ScaleHeight     =   5955
   ScaleWidth      =   8250
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOtherDesc 
      Appearance      =   0  'Flat
      DataField       =   "SC_OTHER_DESC"
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
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "00-Other Description"
      Top             =   4720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtWorkerType 
      Appearance      =   0  'Flat
      DataField       =   "SC_WORKER_TYPE"
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
      Left            =   5520
      TabIndex        =   21
      Top             =   4335
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7080
      Top             =   3840
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
      TabIndex        =   15
      Top             =   5295
      Width           =   8250
      _Version        =   65536
      _ExtentX        =   14552
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
         Left            =   6720
         TabIndex        =   16
         Tag             =   "Select Province listed above"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Tag             =   "Close and exit screen"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1035
         TabIndex        =   6
         Tag             =   "Edit the information above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1845
         TabIndex        =   7
         Tag             =   "Save the changes made"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   8
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3525
         TabIndex        =   9
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4335
         TabIndex        =   10
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5745
         TabIndex        =   17
         Tag             =   "Print the Province listing report"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   5235
      TabIndex        =   13
      Tag             =   "Find specific record"
      Top             =   5610
      Width           =   735
   End
   Begin VB.TextBox txtFindDesc 
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
      Left            =   1080
      TabIndex        =   12
      Tag             =   "00-Search Description"
      Top             =   5640
      Width           =   3975
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
      Left            =   360
      MaxLength       =   4
      TabIndex        =   11
      Tag             =   "00-Search Code"
      Top             =   5640
      Width           =   540
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "PP_COMPNO"
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
      Left            =   6360
      MaxLength       =   3
      TabIndex        =   14
      Text            =   "001"
      Top             =   5040
      Visible         =   0   'False
      Width           =   615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "FsStatusCategory.frx":0000
      Height          =   3105
      Left            =   120
      OleObjectBlob   =   "FsStatusCategory.frx":0014
      TabIndex        =   0
      Tag             =   "Province Listings"
      Top             =   120
      Width           =   8055
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   4440
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SC_EMP"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   1980
      TabIndex        =   1
      Tag             =   "00-Enter Status Code"
      Top             =   3480
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SC_PT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   1980
      TabIndex        =   2
      Tag             =   "00-Category Codes"
      Top             =   3840
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin VB.ComboBox comWorkerType 
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
      ItemData        =   "FsStatusCategory.frx":45F0
      Left            =   2280
      List            =   "FsStatusCategory.frx":45F2
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "10-Type of Worker"
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lblOtherDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Other Description"
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
      Left            =   960
      TabIndex        =   22
      Top             =   4770
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblEEType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Type"
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
      Top             =   4380
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   3525
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   3885
      Width           =   2355
   End
End
Attribute VB_Name = "frmSEmpTypeMatrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew As Integer
Dim fglbNewRec%, xOldCode As String
Dim fglbID As Integer
Dim xLinkItem As String
Dim xStatus, xCategory As String

Private Function chkStatus_Category()
Dim prov$, Msg$
Dim rsStatCat As New ADODB.Recordset
Dim SQLQ As String

chkStatus_Category = False

On Error GoTo chkStatus_Category_Err

If Len(clpCode(0)) < 1 Then
    MsgBox lblTitle(0).Caption & " is a required field"
    clpCode(0).SetFocus
    Exit Function
End If

If Not clpCode(0).ListChecker Then
    Exit Function
End If
If Not clpCode(1).ListChecker Then
    Exit Function
End If

If Len(clpCode(1)) < 1 Then
    MsgBox lblTitle(1).Caption & " is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

'Check if this combination of Employment Status and Category already exists
Set rsStatCat = Nothing
SQLQ = "SELECT * FROM HR_EMPLOYEE_MATRIX"
SQLQ = SQLQ & " WHERE SC_EMP = '" & clpCode(0).Text & "'"
SQLQ = SQLQ & " AND SC_PT = '" & clpCode(1).Text & "'"
rsStatCat.Open SQLQ, gdbAdoIhr001, adOpenStatic
If fglbNewRec% = True Then
    If Not rsStatCat.EOF Then
        'Combination already exist
        MsgBox "This Employee Type matrix already exists."
        clpCode(0).SetFocus
        Exit Function
    End If
    rsStatCat.Close
Else
    If (xStatus <> clpCode(0).Text) Or (xCategory <> clpCode(1).Text) Then
        If Not rsStatCat.EOF Then
            'Combination already exist
            MsgBox "This Employee Type matrix already exists."
            clpCode(0).SetFocus
            Exit Function
        End If
        rsStatCat.Close
    End If
End If

If comWorkerType.ListIndex = -1 Then
    MsgBox "Employee Type is a required field"
    comWorkerType.SetFocus
    Exit Function
End If

If comWorkerType.ListIndex = 12 And Len(Trim(txtOtherDesc.Text)) = 0 Then
    MsgBox "If Employee Type is 'Other' then 'Other Description' cannot be left blank"
    txtOtherDesc.SetFocus
    Exit Function
End If

chkStatus_Category = True

Exit Function

chkStatus_Category_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HR_EMPLOYEE_MATRIX", "chkStatus_Category")
Resume Next

End Function

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

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
xStatus = ""
xCategory = ""
comWorkerType.ListIndex = -1

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_EMPLOYEE_MATRIX", "Cancel")
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

'Private Sub cmdFind_Click()
'Dim SQLQ$
'
'txtFindKey.SetFocus 'added by Marlon Cowan 9/16/97
'
'If Len(txtFindKey) > 0 Then
'    SQLQ$ = "CODE >= '" & txtFindKey.Text & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ$
'    If Data1.Recordset.EOF Then
'        Data1.Refresh
'    Else
'        txtFindKey = ""
'    End If
'    Exit Sub
'End If
'
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
'
'End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
xStatus = clpCode(0).Text
xCategory = clpCode(1).Text

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
xStatus = ""
xCategory = ""
comWorkerType.ListIndex = -1

fglbNewRec% = True

Call modSTUPD(True)

clpCode(0).SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_EMPLOYEE_MATRIX", "AddNew")
Resume Next

End Sub

Private Sub cmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim Desc As String
Dim ProvCode

On Error GoTo OK_Err

If Not chkStatus_Category() Then Exit Sub

Data1.Recordset("SC_COMPNO") = txtComp
If txtWorkerType.Text = "OTHR" Then
    Data1.Recordset("SC_OTHER_DESC") = txtOtherDesc.Text
End If
Data1.Recordset("SC_LDATE") = Format(Now, "SHORT DATE")
Data1.Recordset("SC_LTIME") = Time$
Data1.Recordset("SC_LUSER") = glbUserID
Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("SC_ID")
    Data1.Refresh
    Data1.Recordset.Find "SC_ID=" & fglbID & " "
End If

fglbNewRec% = False

Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption

Resume Next
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_EMPLOYEE_MATRIX", "Update")
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

RHeading = "Employee Type Matrix"
Me.vbxCrystal.WindowTitle = "Employee Type Matrix Report"
Me.vbxCrystal.BoundReportHeading = "Employee Type Matrix"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub comWorkerType_Click()
    If comWorkerType.ListIndex <> -1 Then
        txtOtherDesc.Visible = False
        lblOtherDesc.Visible = False
        
        Select Case comWorkerType.ListIndex
            Case 0: txtWorkerType.Text = "PFT": txtOtherDesc.Text = ""   '"Permanent Full Time"
            Case 1: txtWorkerType.Text = "PPT": txtOtherDesc.Text = ""    '"Permanent Part Time"
            Case 2: txtWorkerType.Text = "TFT": txtOtherDesc.Text = ""    '"Temporary Full Time"
            Case 3: txtWorkerType.Text = "TPT": txtOtherDesc.Text = ""    '"Temporary Part Time"
            Case 4: txtWorkerType.Text = "CI": txtOtherDesc.Text = ""    '"Casual / Irregular"
            Case 5: txtWorkerType.Text = "SEAS": txtOtherDesc.Text = ""    '"Seasonal"
            Case 6: txtWorkerType.Text = "CONT": txtOtherDesc.Text = ""    '"Contract"
            Case 7: txtWorkerType.Text = "STUD": txtOtherDesc.Text = ""    '"Student"
            Case 8: txtWorkerType.Text = "UT": txtOtherDesc.Text = ""    '"Unpaid / Trainee"
            Case 9: txtWorkerType.Text = "RA": txtOtherDesc.Text = ""    '"Registered Apprentice"
            Case 10: txtWorkerType.Text = "OI": txtOtherDesc.Text = ""    '"Optional Insurance"
            Case 11: txtWorkerType.Text = "OOSC": txtOtherDesc.Text = ""   '"Owner Operator or (Sub)Contractor"
            Case 12
                txtWorkerType.Text = "OTHR"   '"Other"
                lblOtherDesc.Visible = True
                txtOtherDesc.Visible = True
                If Not IsNull(Data1.Recordset("SC_OTHER_DESC")) Then
                    txtOtherDesc.Text = Data1.Recordset("SC_OTHER_DESC")
                End If
        End Select
    End If
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
    
    comWorkerType.Clear
    comWorkerType.AddItem "Permanent Full Time"
    comWorkerType.AddItem "Permanent Part Time"
    comWorkerType.AddItem "Temporary Full Time"
    comWorkerType.AddItem "Temporary Part Time"
    comWorkerType.AddItem "Casual / Irregular"
    comWorkerType.AddItem "Seasonal"
    comWorkerType.AddItem "Contract"
    comWorkerType.AddItem "Student"
    comWorkerType.AddItem "Unpaid / Trainee"
    comWorkerType.AddItem "Registered Apprentice"
    comWorkerType.AddItem "Optional Insurance"
    comWorkerType.AddItem "Owner Operator or (Sub)Contractor"
    comWorkerType.AddItem "Other"
    comWorkerType.ListIndex = -1
    
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT * FROM HR_EMPLOYEE_MATRIX ORDER BY SC_EMP,SC_PT "
    Data1.Refresh

    Call setCaption(lblTitle(1))
    
    
    Call modSTUPD(False)            'Jaddy 10/18/99
    
    Call INI_Controls(Me)
    
    Screen.MousePointer = DEFAULT   '
                                
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
If gSec_Upd_Basic Then  '
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
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
comWorkerType.Enabled = TF
txtOtherDesc.Enabled = TF

txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
vbxTrueGrid.Enabled = FT

'If glbDivInhSel Then
'    cmdSelect.Enabled = False
'End If

End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtWorkerType_Change()
        txtOtherDesc.Visible = False
        lblOtherDesc.Visible = False
        
        Select Case txtWorkerType.Text
            Case "PFT": comWorkerType.ListIndex = 0: txtOtherDesc.Text = ""   '"Permanent Full Time"
            Case "PPT": comWorkerType.ListIndex = 1: txtOtherDesc.Text = ""    '"Permanent Part Time"
            Case "TFT": comWorkerType.ListIndex = 2: txtOtherDesc.Text = ""    '"Temporary Full Time"
            Case "TPT": comWorkerType.ListIndex = 3: txtOtherDesc.Text = ""    '"Temporary Part Time"
            Case "CI": comWorkerType.ListIndex = 4: txtOtherDesc.Text = ""    '"Casual / Irregular"
            Case "SEAS": comWorkerType.ListIndex = 5: txtOtherDesc.Text = ""    '"Seasonal"
            Case "CONT": comWorkerType.ListIndex = 6: txtOtherDesc.Text = ""    '"Contract"
            Case "STUD": comWorkerType.ListIndex = 7: txtOtherDesc.Text = ""    '"Student"
            Case "UT": comWorkerType.ListIndex = 8: txtOtherDesc.Text = ""    '"Unpaid / Trainee"
            Case "RA": comWorkerType.ListIndex = 9: txtOtherDesc.Text = ""    '"Registered Apprentice"
            Case "OI": comWorkerType.ListIndex = 10: txtOtherDesc.Text = ""    '"Optional Insurance"
            Case "OOSC": comWorkerType.ListIndex = 11: txtOtherDesc.Text = ""   '"Owner Operator or (Sub)Contractor"
            Case "OTHR"
                comWorkerType.ListIndex = 12   '"Other"
                lblOtherDesc.Visible = True
                txtOtherDesc.Visible = True
                If Not IsNull(Data1.Recordset("SC_OTHER_DESC")) Then
                    txtOtherDesc.Text = Data1.Recordset("SC_OTHER_DESC")
                End If
            Case Else
                comWorkerType.ListIndex = -1
        End Select
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
                   
    SQLQ = "SELECT * FROM HR_EMPLOYEE_MATRIX "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

    Data1.RecordSource = SQLQ
    Data1.Refresh

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

