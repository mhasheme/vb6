VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmDEPTS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departments"
   ClientHeight    =   6450
   ClientLeft      =   1320
   ClientTop       =   660
   ClientWidth     =   8805
   ControlBox      =   0   'False
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
   ScaleHeight     =   6450
   ScaleWidth      =   8805
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkInactiveCode 
      Alignment       =   1  'Right Justify
      Caption         =   "Inactive Code"
      DataField       =   "DF_INACTIVE"
      DataSource      =   "Data1"
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
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   1395
   End
   Begin INFOHR_Controls.CodeLookup clpLgrCode 
      DataField       =   "DF_GLNO"
      Height          =   285
      Left            =   1410
      TabIndex        =   3
      Tag             =   "00-General Ledger Number"
      Top             =   4410
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4440
      Top             =   4560
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
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
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   330
      Left            =   6000
      TabIndex        =   6
      Tag             =   "Find specific record"
      Top             =   5000
      Width           =   950
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
      Left            =   1800
      TabIndex        =   5
      Tag             =   "00-Search Description"
      Top             =   5040
      Width           =   4005
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
      Left            =   120
      MaxLength       =   7
      TabIndex        =   4
      Tag             =   "00-Department search"
      Top             =   5040
      Width           =   1545
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "DF_CO"
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
      Left            =   6600
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4500
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataField       =   "DF_NAME"
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
      Left            =   4260
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "01-Description of Code"
      Top             =   4080
      Width           =   3915
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      DataField       =   "DF_NBR"
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
      Left            =   1730
      MaxLength       =   7
      TabIndex        =   1
      Tag             =   "01-Department's Code"
      Top             =   4080
      Width           =   1320
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DF_LDATE"
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
      Height          =   315
      Index           =   0
      Left            =   0
      MaxLength       =   25
      TabIndex        =   8
      Text            =   "Ldate"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DF_LTIME"
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
      Height          =   315
      Index           =   1
      Left            =   1680
      MaxLength       =   25
      TabIndex        =   9
      Text            =   "LTime"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DF_LUSER"
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
      Height          =   315
      Index           =   2
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   10
      Text            =   "LUser"
      Top             =   3720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxdepts.frx":0000
      Height          =   3795
      Left            =   0
      OleObjectBlob   =   "fxdepts.frx":0014
      TabIndex        =   0
      Tag             =   "Department Listings"
      Top             =   0
      Width           =   8535
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   5790
      Width           =   8805
      _Version        =   65536
      _ExtentX        =   15531
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
      Begin VB.CommandButton cmdUptGL 
         Appearance      =   0  'Flat
         Caption         =   "Update G/L"
         Height          =   375
         Left            =   7080
         TabIndex        =   24
         Tag             =   "Save changes made"
         Top             =   150
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   60
         TabIndex        =   16
         Tag             =   "Select this Department"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   930
         TabIndex        =   17
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1830
         TabIndex        =   18
         Tag             =   "Edit the information "
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2700
         TabIndex        =   19
         Tag             =   "Save changes made"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3570
         TabIndex        =   20
         Tag             =   "Cancel changes made"
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4500
         TabIndex        =   21
         Tag             =   "Create a new Department"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5340
         TabIndex        =   22
         Tag             =   "Delete Department Listed"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   6180
         TabIndex        =   23
         Tag             =   "Print Departmental Listing"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   1860
         Top             =   150
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
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "G/L"
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
      TabIndex        =   13
      Top             =   4440
      Width           =   285
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   3420
      TabIndex        =   12
      Top             =   4110
      Width           =   495
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   4080
      Width           =   990
   End
End
Attribute VB_Name = "frmDEPTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewRec%
Dim RSDATA As New ADODB.Recordset ' Sam add july 02 * Remove Ado
Dim FRS As ADODB.Recordset
Dim xOldGL As String

Private Function chkDept()
Dim Dept As String, SQLQ As String, Msg$
Dim snapDepts As New ADODB.Recordset

chkDept = False
On Error GoTo chkDept_Err

If Len(txtNumber) < 1 Then
    MsgBox "Department Number is a required field"
    txtNumber.SetFocus
    Exit Function
End If

If Len(txtName) < 1 Then
    MsgBox "Department Description is a required field"
    txtName.SetFocus
    Exit Function
End If

If Len(clpLgrCode.Text) > 0 And clpLgrCode.Caption = "Unassigned" Then
    MsgBox lStr("G/L Number must be valid")
     clpLgrCode.Text = ""
     clpLgrCode.SetFocus
    Exit Function
End If

If fglbNewRec Then
    Dept = CStr(txtNumber)
    SQLQ = "SELECT DF_NBR FROM HRDEPT "
    SQLQ = SQLQ & "WHERE DF_NBR = '" & Dept & "'"
    
    If snapDepts.State <> 0 Then snapDepts.Close
    snapDepts.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapDepts.BOF And snapDepts.EOF Then
        snapDepts.Close
    Else
        Msg$ = "This Department already exists"
        MsgBox Msg$
        snapDepts.Close
        Exit Function
    End If
End If

chkDept = True

Exit Function

chkDept_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkDept", "HRDEPT", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmdCancel_Click()
On Error GoTo Can_Err

RSDATA.CancelUpdate

Call Display_Value

Call modSTUPD(False)    ' reset screen's attributes

cmdClose.Enabled = True
cmdClose.SetFocus

fglbNewRec% = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRDEPT", "Cancel")
Call RollBack '08June99

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()

glbDept = ""
glbDeptDesc = ""

Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Dept As String, SQLQ As String, Msg$, a%
Dim snapEEDepts As New ADODB.Recordset

On Error GoTo DelErr

If Len(txtNumber) < 1 Then Exit Sub
Dept = CStr(txtNumber)

If Dept = "ALL" Or Dept = "All" Then
    MsgBox "You can not delete Department ALL"
    Exit Sub
End If

'Add by Frank 4/23/2001 Begin
Screen.MousePointer = HOURGLASS
cmdDelete.Enabled = False
'Add by Frank 4/23/2001 End

SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_DEPTNO FROM HREMP "
SQLQ = SQLQ & "WHERE ED_DEPTNO = '" & Dept & "'"

If snapEEDepts.State <> 0 Then snapEEDepts.Close
snapEEDepts.Open SQLQ, gdbAdoIhr001, adOpenStatic


If snapEEDepts.BOF And snapEEDepts.EOF Then
    GoTo Lok
Else
    Msg$ = "Employee presently assigned to this department"
    Msg$ = Msg$ & Chr(10) & CStr(snapEEDepts("ED_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & snapEEDepts("ED_SURNAME")
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    snapEEDepts.Close
    GoTo End_line
End If

Lok:    'looks ok to me
snapEEDepts.Close

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

If Not glbCompSerial = "S/N - 2394W" And Not glbCompSerial = "S/N - 2390W" Then
    'St. John    #14796
    'Collectcorp #16311
    Call Codes_Master_Integration("DEPT", txtNumber, , True)
End If

gdbAdoIhr001.BeginTrans
RSDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True


End_line:
cmdDelete.Enabled = True
Screen.MousePointer = DEFAULT
Exit Sub                         '

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRDEPT", "Delete")
Call RollBack '08June99

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "DF_NBR = '" & txtFindKey.Text & "'"
    Data1.Recordset.MoveFirst
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
    SQLQ = "DF_NAME >= '" & txtFindDesc.Text & "'"
    Data1.Recordset.MoveFirst
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
    
txtFindDesc.Enabled = True
txtFindKey.Enabled = True
txtFindKey.SetFocus

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo UpdErr
fglbNewRec% = False
Call modSTUPD(True)
txtNumber.Enabled = False
txtName.Enabled = True
txtName.SetFocus

'Data1.Recordset.Edit
    
Exit Sub

UpdErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpd", "HRPROV", "Refresh")
Call RollBack '08June99

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

chkInactiveCode.Value = 0
txtNumber.Enabled = True
txtNumber.SetFocus

fglbNewRec% = True

Call Set_Control("B", Me)
RSDATA.AddNew

txtComp.Text = "001"

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRDEPT", "AddNew")
Call RollBack '08June99

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim DeptNbr
On Error GoTo OK_Err

If Not chkDept() Then Exit Sub

Call UpdUStats(Me)
DeptNbr = txtNumber

gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, RSDATA)
RSDATA.Update
gdbAdoIhr001.CommitTrans

If glbWFC Then '#28637 Franks 05/18/2016
    If Not xOldGL = clpLgrCode.Text Then
        If Len(clpLgrCode.Text) > 0 Then
            Call UptGLBaseOnDept(txtNumber.Text, clpLgrCode.Text)
        End If
    End If
End If

Data1.RecordSource = "SELECT * FROM HRDEPT ORDER BY DF_INACTIVE, DF_NAME"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Data1.Recordset.Find "DF_NBR='" & DeptNbr & "'"


fglbNewRec% = False

Call modSTUPD(False)

If Not glbCompSerial = "S/N - 2394W" And Not glbCompSerial = "S/N - 2390W" Then
    'St. John    #14796
    'Collectcorp #16311
    Call Codes_Master_Integration("DEPT", txtNumber)
End If

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRDEPT", "Update")
Call RollBack '08June99

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String, xReport

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lStr("Department Listing Report")
Me.vbxCrystal.WindowTitle = RHeading
Me.vbxCrystal.Formulas(0) = "PgHeading = '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"

xReport = glbIHRREPORTS & "rgdept.rpt"

Me.vbxCrystal.ReportFileName = xReport
Me.vbxCrystal.Formulas(1) = "lblDept='" & lStr("Department") & "'"
Me.vbxCrystal.Formulas(2) = "lblGL='" & lStr("G/L#") & "'"
'If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
'Else
'    Me.vbxCrystal.Connect = "PWD=petman;"
'    Me.vbxCrystal.DataFiles(0) = glbIHRDB
'End If

Me.vbxCrystal.Action = 1


End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()

glbDept = Data1.Recordset("DF_NBR")
glbDeptDesc = Data1.Recordset("DF_NAME")

If clpLgrCode.Text = "" Then
    glbGLNum = ""
Else
    glbGLNum = Data1.Recordset("DF_GLNO")
End If

Unload Me

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdUptGL_Click() 'Ticket #28637 Franks 05/18/2016
Dim SQLQ As String, Msg$, a%


Msg = "This function will update employee G/L with '" & clpLgrCode.Text & "' for Department '" & txtNumber.Text & "' "
Msg = Msg & Chr(10) & Chr(10) & "Are You Sure You Want To Do It?"
a% = MsgBox(Msg, 36, "Confirm Update")
If a% <> 6 Then Exit Sub

cmdUptGL.Enabled = False
Call UptGLBaseOnDept(txtNumber.Text, clpLgrCode.Text)
cmdUptGL.Enabled = True

MsgBox "   Finished.   "
End Sub

Private Sub Form_Activate()
Data1.RecordSource = "SELECT * FROM HRDEPT WHERE " & glbSeleDept & " ORDER BY DF_INACTIVE, DF_NAME"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub Form_Load()
Dim SQLQ As String

glbOnTop = "FRMDEPTS"
'Data1.DatabaseName = glbIHRDB

SQLQ = "UPDATE HRDEPT SET DF_INACTIVE = 0 WHERE DF_INACTIVE IS NULL"
gdbAdoIhr001.Execute SQLQ

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HRDEPT WHERE " & glbSeleDept & " ORDER BY DF_INACTIVE, DF_NAME"
Data1.Refresh

Set FRS = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

Screen.MousePointer = HOURGLASS
'Me.vbxTrueGrid.Refresh

Screen.MousePointer = DEFAULT

Call modSTUPD(False)
Call setCaption(lblTitle(0))
Call setCaption(lblTitle(1))
Call setCaption(Me)
Call setCaption(Me.vbxTrueGrid.Columns(0))
Call setCaption(Me.vbxTrueGrid.Columns(1))
Call setCaption(Me.vbxTrueGrid.Columns(2))

If glbVadim Then txtNumber.MaxLength = 4

If Not gSec_Upd_Departments Then       'May99 js
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If

Call INI_Controls(Me)

If glbWFC Then  'Ticket #28637 Franks 05/18/2016
'    cmdUptGL.Visible = True
End If
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

cmdModify.Enabled = FT      'May99 js
cmdFind.Enabled = FT        '
cmdDelete.Enabled = FT      '
cmdNew.Enabled = FT         '
cmdCancel.Enabled = TF      '
cmdOK.Enabled = TF          '
vbxTrueGrid.Enabled = FT
txtFindDesc.Enabled = FT    '
txtFindKey.Enabled = FT     '
clpLgrCode.Enabled = TF     '
txtName.Enabled = TF        '
txtNumber.Enabled = TF      '
chkInactiveCode.Enabled = TF
cmdClose.Enabled = FT       '
cmdSelect.Enabled = FT      '
cmdPrint.Enabled = FT       '
If glbDeptInhSel Then
    cmdSelect.Enabled = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmDEPTS = Nothing  'carmen may 2000
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
'Private Sub txtLgrCode_Change()
'    Dim I%
'    Call LGR_Desc(I)
'End Sub
'Private Sub txtLgrCode_DblClick()   'May99 js
'Dim OLgr As String, OLgrD As String
'OLgr = txtLgrCode.Text
'OLgrD = lblLgrDesc.Caption
'Load frmLEDGER
'frmLEDGER.Show 1
'If Len(glbLgr) < 1 Then
'    txtLgrCode.Text = OLgr
'    lblLgrDesc.Caption = OLgrD
'    lblLgrDesc.Visible = False
'Else
'    txtLgrCode.Text = glbLgr
'    lblLgrDesc.Caption = glbLgrDesc
'    lblLgrDesc.Visible = True
'End If
'End Sub
'Private Sub txtLgrCode_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtLgrCode_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub

Private Sub txtName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr$(KeyAscii)) 'Frank 5/4/2000 Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtNumber_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtNumber_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_DblClick()

If cmdSelect.Enabled Then
    If Not Me.vbxTrueGrid.EditActive Then
        glbDept = Data1.Recordset("DF_NBR")
        glbDeptDesc = Data1.Recordset("DF_NAME")
        Unload Me
    Else
        MsgBox "Save/cancel changes first"
    End If
End If

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    If Not fglbNewRec% Then
        FRS.Requery
        FRS.Bookmark = Bookmark
        If FRS("DF_INACTIVE") Then
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
    
    SQLQ = "SELECT * FROM HRDEPT WHERE " & glbSeleDept
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

Private Function RollBack()
On Error GoTo RR
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
RR:
End Function
''' Sam add July 2002 * Remove ADO
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        
        If RSDATA.State <> 0 Then RSDATA.Close
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HRDEPT WHERE DF_NBR='" & Data1.Recordset!DF_NBR & "'"
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    If IsNull(RSDATA("DF_GLNO")) Then xOldGL = "" Else xOldGL = RSDATA("DF_GLNO")
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'''Sam add July 02 * Remove ADO
Call Display_Value
End Sub

Private Sub UptGLBaseOnDept(xDept, xGL) 'Ticket #28637 Franks 05/18/2016
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xOldGL
    If Len(xDept) = 0 Then Exit Sub
    If Len(xGL) = 0 Then Exit Sub
    
    xGL = Trim(xGL)
    
    Screen.MousePointer = HOURGLASS

    SQLQ = "SELECT * FROM HREMP WHERE ED_DEPTNO = '" & xDept & "' "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsEmp.EOF
        If IsNull(rsEmp("ED_GLNO")) Then xOldGL = "" Else xOldGL = rsEmp("ED_GLNO")
        If Not xOldGL = xGL Then
            'update Audit
            If Len(xGL) > 0 Then
                Call WFC_AUDIT_ByField(rsEmp("ED_EMPNBR"), "M", "ED_GLNO", xGL, xOldGL, rsEmp)
            End If
            
            If Len(xGL) = 0 Then
                rsEmp("ED_GLNO") = Null
            Else
                rsEmp("ED_GLNO") = Left(xGL, 25)
            End If
            rsEmp.Update
        End If
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    
    Screen.MousePointer = DEFAULT
End Sub

Private Function WFC_AUDIT_ByField(xEmpNo, ACTX, xFieldName, xNewVal, xOldVal, rsEmp As ADODB.Recordset)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
Dim SQLQ
    
'''On Error GoTo AUDIT_ERR
WFC_AUDIT_ByField = False

'rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo, gdbAdoIhr001, adOpenKeyset

If Not rsEmp.EOF Then
    If IsNull(rsEmp("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsEmp("ED_PT")
    End If
    If IsNull(rsEmp("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsEmp("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If

strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_PER, AU_BAMT, AU_UNITCOST,AU_CEASEDATE, "
strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS,AU_DEPT_GL,AU_OLD_GL "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

xADD = False

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")

If xFieldName = "ED_GLNO" Then
    rsTA("AU_DEPT_GL") = IIf(xNewVal = "", Null, xNewVal)
    rsTA("AU_OLD_GL") = IIf(xOldVal = "", Null, xOldVal)
End If

rsTA("AU_LUSER") = glbUserID
rsTA("AU_LDATE") = Date
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update


End_line:

WFC_AUDIT_ByField = True
Exit Function
AUDIT_ERR:

End Function

