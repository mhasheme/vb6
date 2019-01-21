VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmProjectCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Code Master"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkInactiveCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Inactive Code"
      DataField       =   "PC_INACTIVE"
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   1395
   End
   Begin VB.TextBox txtShortDesc 
      Appearance      =   0  'Flat
      DataField       =   "GL_NUMBER"
      Height          =   285
      Left            =   1590
      MaxLength       =   10
      TabIndex        =   3
      Tag             =   "01-Short Description"
      Top             =   4590
      Width           =   2295
   End
   Begin VB.CheckBox chkSplit 
      Alignment       =   1  'Right Justify
      Caption         =   "Split"
      DataField       =   "SPLIT"
      Height          =   285
      Left            =   5550
      TabIndex        =   4
      Top             =   4200
      Width           =   795
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3480
      MaxLength       =   25
      TabIndex        =   10
      Text            =   "LUser"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   9
      Text            =   "LTime"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   120
      MaxLength       =   25
      TabIndex        =   8
      Text            =   "Ldate"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "PROJECT_CODE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   11
      TabIndex        =   1
      Tag             =   "01-Project Code"
      Top             =   4200
      Width           =   1365
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "DESCRIPTION"
      Height          =   285
      Left            =   1590
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "01-Description of Code"
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   11
      TabIndex        =   5
      Tag             =   "00-Search Project Code"
      Top             =   4980
      Width           =   1350
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1590
      MaxLength       =   30
      TabIndex        =   6
      Tag             =   "00-Search Description"
      Top             =   4980
      Width           =   3735
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5610
      TabIndex        =   7
      Tag             =   "Find specific record"
      Top             =   4950
      Width           =   720
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5310
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      TabIndex        =   11
      Top             =   5865
      Width           =   7455
      _Version        =   65536
      _ExtentX        =   13150
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   19
         Tag             =   "Print Project Code Listing"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5070
         TabIndex        =   18
         Tag             =   "Delete Project Code listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4260
         TabIndex        =   17
         Tag             =   "Create a new Project Code "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         TabIndex        =   16
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Tag             =   "Save changes made"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Tag             =   "Edit the information "
         Top             =   90
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         TabIndex        =   13
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   15
         TabIndex        =   12
         Tag             =   "Select this Project Code "
         Top             =   105
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   1935
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmProjectCode.frx":0000
      Height          =   3795
      Left            =   30
      OleObjectBlob   =   "frmProjectCode.frx":0014
      TabIndex        =   0
      Tag             =   "Project code Listings"
      Top             =   0
      Width           =   7395
   End
   Begin VB.Label lblShortDesc 
      AutoSize        =   -1  'True
      Caption         =   "Short Description"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   4620
      Width           =   1215
   End
End
Attribute VB_Name = "frmProjectCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNoRecords%
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO

'Hemu
Dim fglbMultiSelect As Boolean
'Hemu
Dim FRS As ADODB.Recordset

Private Function chkChargeCode()
Dim xCode As String, SQLQ As String, msg$
Dim rsChargeCode As New ADODB.Recordset

chkChargeCode = False
On Error GoTo chkChargeCode_Err

If Len(txtCode) < 1 Then
    MsgBox lStr("Project Code is a required field")
    txtCode.SetFocus
    Exit Function
End If

If Len(txtDesc) < 1 Then
    MsgBox lStr("Project Code Description is a required field")
    txtDesc.SetFocus
    Exit Function
End If


If fglbNewRec% Then
    xCode = CStr(txtCode)
    SQLQ = "SELECT PROJECT_CODE from HR_PROJECT_CODE "
    SQLQ = SQLQ & "WHERE PROJECT_CODE = '" & xCode & "'"
    
    If rsChargeCode.State <> 0 Then rsChargeCode.Close
    rsChargeCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If rsChargeCode.BOF And rsChargeCode.EOF Then
        rsChargeCode.Close
    Else
        msg$ = lStr("This Project Code already exists")
        MsgBox msg$
        rsChargeCode.Close
        Exit Function
    End If
End If

chkChargeCode = True

Exit Function

chkChargeCode_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkChargeCode", "HR_PROJECT_CODE", "Cancel")
Resume Next

End Function

Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

'Data1.UpdateControls    ' returns without saving
rsDATA.CancelUpdate

Call Set_Control("R", Me, rsDATA)
'Data1.Recordset.CancelUpdate
'If Not glbSQL Then Call Pause(0.5)
'Data1.Refresh
Call modSTUPD(False)  ' reset screen's attributes

cmdClose.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPROv", "Cancel")
'Resume Next

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()

glbCode = ""

Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim xCode As String, SQLQ As String, msg$, a%
'Dim rsChargeCode As New ADODB.Recordset

On Error GoTo DelErr

If Len(txtCode) < 1 Then Exit Sub

xCode$ = CStr(txtCode)

If Data1.Recordset.RecordCount = 1 Then
    MsgBox lStr("You can not delete the last Project Code.")
    Exit Sub
End If

'SQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE "
'SQLQ = SQLQ & " WHERE AD_CHRGCODE= '" & xCode & "'"
'SQLQ = SQLQ & " GROUP BY AD_EMPNBR "
'
'rsChargeCode.Open SQLQ, gdbADOIHR001, adOpenStatic
'
'If rsChargeCode.BOF And rsChargeCode.EOF Then
'    GoTo Lok
'Else
'    Msg$ = lStr("Employee presently assigned to this Project Code")
'    Msg$ = Msg$ & Chr(10) & ShowEmpnbr(rsChargeCode("ED_EMPNBR"))
'   ' Msg$ = Msg$ & Chr(10) & rsChargeCode("ED_SURNAME")
'    Msg$ = Msg$ & Chr(10) & "Delete aborted."
'    MsgBox Msg$
'    rsChargeCode.Close
'    Exit Sub
'End If
'
'Lok:    'looks ok to me
'rsChargeCode.Close

msg = "Are You Sure You Want To Delete "
msg = msg & "This Record?"
a% = MsgBox(msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

'Data1.Recordset.Delete
'If Not glbSQL Then Call Pause(0.5)
'Data1.Refresh
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_PROJECT_CODE", "Delete")
Resume Next

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "PROJECT_CODE = '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
        
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
        
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ = "DESCRIPTION >= '" & txtFindDesc.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    
        Set FRS = Data1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True
    
    Else
        txtFindDesc = ""
    End If
    Exit Sub
End If

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)

txtCode.Enabled = False
txtDesc.Enabled = True
chkSplit.Enabled = True
txtShortDesc.Enabled = True
txtDesc.SetFocus

'Data1.Recordset.Edit

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Unload Me
End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

'glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True
txtCode.Enabled = True
chkInactiveCode.Value = 0
txtCode.SetFocus

'Data1.Recordset.AddNew
Call Set_Control("B", Me)
rsDATA.AddNew

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_PROJECT_CODE", "AddNew")
Resume Next

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim xCode
On Error GoTo OK_Err

If Not chkChargeCode() Then Exit Sub

Call UpdUStats(Me)
xCode = txtCode


gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, rsDATA)
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.RecordSource = "SELECT * FROM HR_PROJECT_CODE ORDER BY PC_INACTIVE, PROJECT_CODE"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True


Data1.Recordset.Find "PROJECT_CODE='" & xCode & " '"
fglbNewRec% = False

Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PROJECT_CODE", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String, xReport

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    RHeading = lStr("Project Code") & " Listing Report"
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading

    xReport = glbIHRREPORTS & "rgProjCode.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.Formulas(0) = "lblGL='" & lStr("G/L #") & "'"
    Me.vbxCrystal.Connect = RptODBC_SQL

    Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()
Dim X
If fglbMultiSelect And vbxTrueGrid.SelBookmarks.count <> 0 Then
    If vbxTrueGrid.SelBookmarks.count > 1000 Then
        MsgBox vbxTrueGrid.SelBookmarks.count & " codes are selected" + Chr(10) + " Please make that less than 1000 codes"
        Exit Sub
    End If
    glbCode = ""
    For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
        vbxTrueGrid.Bookmark = vbxTrueGrid.SelBookmarks(X)
        glbCode = glbCode & Data1.Recordset!PROJECT_CODE & ","
    Next
    glbCode = Left(glbCode, Len(glbCode) - 1)
Else
    glbCode = Data1.Recordset("PROJECT_CODE")
End If

Unload Me

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_PROJECT_CODE", "SELECT")

End Sub

Private Sub Form_Activate()
Dim xStr

Data1.RecordSource = "SELECT * FROM HR_PROJECT_CODE ORDER BY PC_INACTIVE, PROJECT_CODE"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

If fglbMultiSelect Then
    vbxTrueGrid.MultiSelect = 2
    If glbCode <> "" Then
        With Data1.Recordset
            If Not .EOF Then .MoveLast
            xStr = glbCode & ","
            Do Until .BOF
                If InStr(glbCode & ",", !PROJECT_CODE & ",") <> 0 Then
                    xStr = Replace(xStr, !PROJECT_CODE & ",", "")
                    vbxTrueGrid.SelBookmarks.Add vbxTrueGrid.Bookmark
                    DoEvents
                    If Trim(xStr) = "" Then Exit Do
                End If
                .MovePrevious
            Loop
        End With
    End If
Else
    vbxTrueGrid.MultiSelect = 1
End If

End Sub

Private Sub Form_Load()
Dim SQLQ

SQLQ = "UPDATE HR_PROJECT_CODE SET PC_INACTIVE = 0 WHERE PC_INACTIVE IS NULL"
gdbAdoIhr001.Execute SQLQ

SQLQ = "SELECT * FROM HR_PROJECT_CODE "
If glbOracle Then
    SQLQ = SQLQ & " ORDER BY PC_INACTIVE, UPPER(PROJECT_CODE)"
Else
    SQLQ = SQLQ & " ORDER BY PC_INACTIVE, PROJECT_CODE"
End If

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Me.vbxTrueGrid.Columns(2).Caption = lStr(Me.vbxTrueGrid.Columns(2).Caption)
Me.vbxTrueGrid.Refresh

Screen.MousePointer = DEFAULT

If glbCompSerial <> "S/N - 2192W" Then ' For county of essex
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(3).Visible = False
    chkSplit.Visible = False
    lblShortDesc.Visible = False
    txtShortDesc.Visible = False
End If

Call INI_Controls(Me)
Call modSTUPD(False)

If Not gSec_Upd_Project_Code Then     'May99 js
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If                          '

Call setCaption(Me)
'Call setCaption(Me.vbxTrueGrid.Columns.Item(0))
'Call setCaption(Me.vbxTrueGrid.Columns.Item(1))

End Sub

Private Sub Form_LostFocus()

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

cmdModify.Enabled = FT          'May99 js
cmdFind.Enabled = FT            '
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
                                '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '
                                '
txtCode.Enabled = TF
txtDesc.Enabled = TF            '
chkSplit.Enabled = TF                               '
txtShortDesc.Enabled = TF
chkInactiveCode.Enabled = TF
vbxTrueGrid.Enabled = FT 'Jaddy 11/12/99
txtFindDesc.Enabled = FT        '
txtFindKey.Enabled = FT         '
cmdClose.Enabled = FT           '
cmdSelect.Enabled = False
cmdPrint.Enabled = FT           '
If Data1.Recordset.EOF Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
        
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
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

Private Sub txtDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()
    
If Not Me.vbxTrueGrid.EditActive Then
    glbCode = Data1.Recordset("PROJECT_CODE")
    Unload Me
Else
    MsgBox "Save/cancel changes first"
End If

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If Not fglbNewRec% Then
        FRS.Requery
        FRS.Bookmark = Bookmark
        If FRS("PC_INACTIVE") Then
            RowStyle.ForeColor = vbRed
        End If
    End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
   
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT * FROM HR_PROJECT_CODE "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
    
    Set FRS = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    If Me.vbxTrueGrid.EditActive Then
        cmdOK.SetFocus
    Else
        cmdClose.SetFocus
    End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub

Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
  
    SQLQ = "select * from HR_PROJECT_CODE WHERE PROJECT_CODE='" & Data1.Recordset!PROJECT_CODE & "'"
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub

