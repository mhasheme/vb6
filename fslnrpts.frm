VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSCusRPTs 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Custom Reports Master"
   ClientHeight    =   8490
   ClientLeft      =   525
   ClientTop       =   1470
   ClientWidth     =   11880
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
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkUseHS 
      Caption         =   "Use H&&S Incident Fields for Validation"
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
      Left            =   3120
      TabIndex        =   17
      Top             =   6600
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Frame frmAT 
      Height          =   435
      Left            =   330
      TabIndex        =   10
      Top             =   5280
      Width           =   6015
      Begin VB.OptionButton optAT 
         Caption         =   "Terminated Employee"
         Height          =   255
         Index           =   1
         Left            =   2910
         TabIndex        =   7
         Top             =   150
         Width           =   2175
      End
      Begin VB.OptionButton optAT 
         Caption         =   "Active Employee"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   150
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.CheckBox chkTermination 
         Alignment       =   1  'Right Justify
         Caption         =   "Terminated Employees"
         DataField       =   "RT_TERMINATION"
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
         Left            =   6480
         TabIndex        =   11
         Tag             =   "40-Termination Employees"
         Top             =   180
         Value           =   2  'Grayed
         Visible         =   0   'False
         Width           =   2145
      End
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      DataField       =   "RT_FILENAME"
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
      Left            =   2490
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "00-File Name"
      Top             =   2310
      Width           =   6435
   End
   Begin VB.TextBox txtRPTName 
      Appearance      =   0  'Flat
      DataField       =   "RT_RPTNAME"
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
      Left            =   2490
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "00-Report English Name"
      Top             =   1950
      Width           =   6435
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   2520
      TabIndex        =   4
      Tag             =   "Dir"
      Top             =   3060
      Width           =   2355
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2520
      TabIndex        =   3
      Tag             =   "Drive"
      Top             =   2670
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2625
      Left            =   4980
      Pattern         =   "*.rpt;*.xls"
      TabIndex        =   5
      Tag             =   "File"
      Top             =   2640
      Width           =   3945
   End
   Begin VB.ComboBox cmbTable 
      DataField       =   "RT_DATETABLE"
      Height          =   315
      Left            =   3150
      TabIndex        =   8
      Tag             =   "Table Name"
      Text            =   "cmbTable"
      Top             =   5760
      Width           =   3795
   End
   Begin VB.ComboBox cmbField 
      DataField       =   "RT_DATEFIELD"
      Height          =   315
      Left            =   3150
      TabIndex        =   9
      Tag             =   "Field Name"
      Text            =   "cmbField"
      Top             =   6150
      Width           =   3795
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fslnrpts.frx":0000
      Height          =   1785
      Left            =   210
      OleObjectBlob   =   "fslnrpts.frx":0014
      TabIndex        =   0
      Top             =   90
      Width           =   9810
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   960
      Top             =   7440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowWidth     =   480
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileType   =   2
      BoundReportHeading=   "RGELIST"
      BoundReportFooter=   -1  'True
      WindowControls  =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   405
      Left            =   5520
      Top             =   7800
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   714
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
   Begin VB.Label Label4 
      Caption         =   "Field"
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
      Left            =   2400
      TabIndex        =   16
      Top             =   6210
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Table"
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
      Left            =   2400
      TabIndex        =   15
      Top             =   5850
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Date Range Selection for"
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
      Left            =   360
      TabIndex        =   14
      Top             =   5850
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Name"
      Height          =   195
      Left            =   270
      TabIndex        =   13
      Top             =   2340
      Width           =   855
   End
   Begin VB.Label lblRTPName 
      Caption         =   "Report's English Name"
      Height          =   255
      Left            =   270
      TabIndex        =   12
      Top             =   1950
      Width           =   2535
   End
End
Attribute VB_Name = "frmSCusRPTs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim ORPTName As String
Dim fglbPATH As String
Dim OTABLE As String
Dim OAT As Boolean
Dim fglbNew As Boolean
Private Sub chkTermination_Click()
optAT(0).Value = True
If Not Data1.Recordset.EOF Then
    If chkTermination.Value = 1 Then
        optAT(1).Value = True
    End If
End If
Call cmbDateTable
End Sub


Private Sub cmbField_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbTable_Change()
Call cmbDateField
End Sub

Private Sub cmbTable_Click()
If OTABLE <> cmbTable Then
    Call cmbDateField
End If
End Sub

Private Sub cmbTable_GotFocus()
OTABLE = cmbTable
Call SetPanHelp(ActiveControl)
End Sub


Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
rsDATA.CancelUpdate
Call Display_Value


'Call ST_UPD_MODE(False)  ' reset screen's attributes
'Call ST_UPD_MODE(True)

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRLNRPT", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub

'Private Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Unload Me

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub



Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err


Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub
gdbAdoIhr001.Execute "DELETE FROM HR_SECRPT WHERE " & Field_SQL("FUNCTION") & "='" & txtRPTName & "'"

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If


Me.vbxTrueGrid.SetFocus
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRLNRPT", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

'Sub cmdDelete_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
Dim SQLQ As String

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
ORPTName = txtRPTName
On Error GoTo Edit_Err

'txtRPTName.SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "HRLNRPT", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub



Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True
Call SET_UP_MODE
'Call ST_UPD_MODE(True)


On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

ORPTName = ""
cmbTable = ""
cmbField = ""
optAT(0) = True
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRLNRPT", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x%
Dim xID
On Error GoTo OK_Err
If Not chkCustomRPT Then Exit Sub
If Len(ORPTName) > 0 And ORPTName <> txtRPTName Then Call UPD_Security

gdbAdoIhr001.BeginTrans
rsDATA!RT_COMPNO = "001"
chkTermination.Value = IIf(optAT(1), 1, 0)
Call Set_Control("U", Me, rsDATA)
rsDATA("RT_DATETABLE") = cmbTable
rsDATA("RT_DATEFIELD") = cmbField
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Me.vbxTrueGrid.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRLNRPT", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub


Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False


Me.vbxCrystal.WindowTitle = "Custom Reports Setup Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGCUSRPT.rpt"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub
Sub cmdView_Click()
Dim RHeading As String, xReport, x%

'cmdPrint.Enabled = False

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup


Me.vbxCrystal.WindowTitle = "Custom Reports Setup Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        For x% = 0 To 1
            Me.vbxCrystal.DataFiles(x%) = glbIHRDB
        Next
    End If
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGCUSRPT.rpt"
 Me.vbxCrystal.Destination = 0
 
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub


'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRLNRPT", "SELECT")


End Sub



Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Dir1_LostFocus()
Call Get_File
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1
End Sub

Private Sub Drive1_LostFocus()
Call Get_File
End Sub

Private Sub File1_Click()
Call Get_File
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
glbOnTop = Me.name
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%
Dim SQLQ

Screen.MousePointer = HOURGLASS
Data1.ConnectionString = glbAdoIHRDB

If glbWFC Then 'Ticket #20859 Franks 01/30/2012
    chkUseHS.DataField = "RT_USER_FLAG"
    chkUseHS.Visible = True
End If

If glbOracle Then
    SQLQ = "SELECT RT_DATETABLE || '.' || RT_DATEFIELD AS RT_DATENAME,"
    SQLQ = SQLQ & "RT_COMPNO,RT_ID,RT_FILENAME,RT_RPTNAME,RT_DATETABLE,RT_DATEFIELD,RT_TERMINATION "
    SQLQ = SQLQ & " FROM HR_CUSTOMRPT "
Else
    SQLQ = "SELECT *,RT_DATETABLE + '.' + RT_DATEFIELD AS RT_DATENAME FROM HR_CUSTOMRPT "
End If
Data1.RecordSource = SQLQ
Data1.Refresh



If vbxTrueGrid.Visible Then Me.vbxTrueGrid.SetFocus

Call cmbDateTable
Call cmbDateField
Call Display_Value

Call ST_UPD_MODE(False)
If Not gSec_Upd_CustomReport Then
'    cmdDelete.Enabled = False
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
End If
Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmSCusRPTs = Nothing

End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF


txtRPTName.Enabled = TF
File1.Enabled = TF
Dir1.Enabled = TF
Drive1.Enabled = TF
frmAT.Enabled = TF
cmbTable.Enabled = TF
cmbField.Enabled = TF
'vbxTrueGrid.Enabled = FT
End Sub






Private Sub optAT_GotFocus(Index As Integer)
OAT = optAT(0)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub optAT_LostFocus(Index As Integer)
If OAT <> optAT(0) Then Call cmbDateTable
End Sub

Private Sub txtFileName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtRPTName_GotFocus()
Call SetPanHelp(ActiveControl)
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
        
        If glbOracle Then
            SQLQ = "SELECT RT_DATETABLE || '.' || RT_DATEFIELD AS RT_DATENAME,"
            SQLQ = SQLQ & "RT_COMPNO,RT_ID,RT_FILENAME,RT_RPTNAME,RT_DATETABLE,RT_DATEFIELD,RT_TERMINATION "
            SQLQ = SQLQ & " FROM HR_CUSTOMRPT "
        Else
            SQLQ = "SELECT *,RT_DATETABLE + '.' + RT_DATEFIELD AS RT_DATENAME FROM HR_CUSTOMRPT "
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
'If KeyAscii = 9 Then ' if the tab key was struck
'    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdClose.SetFocus
'    End If
'End If

End Sub


Private Function chkCustomRPT()
Dim xFile
On Error GoTo err_chkCustomRPT

chkCustomRPT = False

If Len(txtRPTName) = 0 Then
    MsgBox "Report's English Name is a required field"
    txtRPTName.SetFocus
    Exit Function
End If
If Len(txtFileName) = 0 Then
    MsgBox "File Name is a required field"
    File1.SetFocus
    Exit Function
Else
    If UCase(glbIHRREPORTS) = UCase(fglbPATH) Then
        xFile = glbIHRREPORTS & txtFileName
    Else
        xFile = txtFileName
    End If
    If Dir(xFile) = "" Then
        MsgBox "File " & txtFileName & " do not exist!"
        File1.SetFocus
        Exit Function
    End If
End If
If Len(cmbTable) > 0 And Len(cmbField) = 0 Then
    MsgBox "If Table Name Entered, Field Name Must be Entered"
    cmbTable.SetFocus
    Exit Function
End If
If Len(cmbTable) = 0 And Len(cmbField) > 0 Then
    MsgBox "If Field Name Entered, Table Name Must be Entered"
    cmbField.SetFocus
    Exit Function
End If

chkCustomRPT = True
Exit Function
err_chkCustomRPT:
If Err = 52 Then
    MsgBox "File " & txtFileName & " do not exist!"
    File1.SetFocus
    Exit Function
End If
End Function

Private Sub Show_Path()
On Error Resume Next
Dim xFile
Dim x

fglbPATH = glbIHRREPORTS
If InStr(txtFileName, ":") <> 0 Then
    fglbPATH = Left(txtFileName, InStrRev(txtFileName, "\"))
    xFile = Mid(txtFileName, InStrRev(txtFileName, "\") + 1)
Else
    xFile = txtFileName
End If
Drive1.Drive = Left(fglbPATH, InStr(glbIHRREPORTS, ":"))
Dir1.Path = fglbPATH
File1.Path = fglbPATH

For x = 0 To File1.ListCount - 1
    If UCase(File1.List(x)) = UCase(xFile) Then
        File1.selected(x) = True
    End If
Next
End Sub
Private Sub Get_File()
fglbPATH = Dir1.Path
fglbPATH = fglbPATH & IIf(Right(fglbPATH, 1) <> "\", "\", "")
If UCase(fglbPATH) = UCase(glbIHRREPORTS) Then
    txtFileName = File1
Else
    txtFileName = fglbPATH & File1
End If
End Sub
Private Sub cmbDateTable()
Dim rsINFO As New ADODB.Recordset
Dim xTable
Dim SQLQ
cmbTable.Clear

SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"

If optAT(1) Then
    SQLQ = SQLQ & " AND TERMINATION_TABLE<>0"
Else
    SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
End If
rsINFO.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly

Do Until rsINFO.EOF
    xTable = rsINFO("TABLE_NAME")
    If xTable <> "" Then cmbTable.AddItem xTable
    rsINFO.MoveNext
Loop
If glbCompSerial = "S/N - 2433W" Then 'Ticket #23379 Franks 03/06/2013 Kerry's Place
    cmbTable.AddItem "WT_ATTEND"
End If

End Sub
Private Sub cmbDateField()
On Error Resume Next
Dim rsTable As New ADODB.Recordset
Dim x
Dim xTable As String
cmbField.Clear
If cmbTable = "" Then Exit Sub
xTable = cmbTable
If optAT(0) Then
    rsTable.Open xTable, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
Else
    rsTable.Open xTable, gdbAdoIhr001X, adOpenStatic, adLockReadOnly, adCmdTableDirect
End If '
With rsTable
    For x = 0 To .Fields.count - 1
        If .Fields(x).Type = adDate Or .Fields(x).Type = adDBDate Or .Fields(x).Type = adDBTime Or .Fields(x).Type = adDBTimeStamp Then
            cmbField.AddItem rsTable.Fields(x).name
        End If
    Next
End With
End Sub

Private Sub UPD_Security()
Dim xRPTNAME, xORPTName, SQLQ
xRPTNAME = Replace(txtRPTName, "'", "''")
xORPTName = Replace(ORPTName, "'", "''")
SQLQ = "UPDATE HR_SECRPT "
SQLQ = SQLQ & " SET " & Field_SQL("FUNCTION") & "='" & xRPTNAME & "'"
SQLQ = SQLQ & " WHERE " & Field_SQL("FUNCTION") & "='" & xORPTName & "'"
gdbAdoIhr001.Execute SQLQ
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

End Sub
Private Sub Display_Value()
Dim SQLQ
cmbTable = ""
cmbField = ""
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    SQLQ = "SELECT * from HR_CUSTOMRPT"
    If glbtermopen Then
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
    SQLQ = "SELECT * from HR_CUSTOMRPT WHERE RT_ID= " & Data1.Recordset!RT_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    If Not IsNull(Data1.Recordset("RT_DATETABLE")) Then cmbTable = rsDATA("RT_DATETABLE")
    If Not IsNull(Data1.Recordset("RT_DATEFIELD")) Then cmbField = rsDATA("RT_DATEFIELD")
End If
Call SET_UP_MODE
Me.cmdModify_Click
Call Show_Path
End Sub


Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property
Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_CustomReport
End Property

Public Property Get Addable() As Boolean

Addable = True
End Property
Public Property Get Updateble() As Boolean

Updateble = True
End Property
Public Property Get Deleteble() As Boolean

Deleteble = True
End Property
Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub
