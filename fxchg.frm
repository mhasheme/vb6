VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCHARGECODE 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charge Code Master"
   ClientHeight    =   5775
   ClientLeft      =   1125
   ClientTop       =   795
   ClientWidth     =   6900
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5775
   ScaleWidth      =   6900
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3480
      MaxLength       =   25
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
      Text            =   "Ldate"
      Top             =   3840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "CHARGE_CODE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   1
      Tag             =   "01-Charge Code"
      Top             =   4200
      Width           =   1065
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "DESCRIPTION"
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "01-Description of Code"
      Top             =   4200
      Width           =   3735
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "00-Search Charge Code"
      Top             =   4775
      Width           =   1080
   End
   Begin VB.TextBox txtFindDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Tag             =   "00-Search Description"
      Top             =   4770
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
      Left            =   5340
      TabIndex        =   5
      Tag             =   "Find specific record"
      Top             =   4710
      Width           =   720
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5040
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
      TabIndex        =   13
      Top             =   5115
      Width           =   6900
      _Version        =   65536
      _ExtentX        =   12171
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
         TabIndex        =   17
         Tag             =   "Print Charge Code Listing"
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
         TabIndex        =   16
         Tag             =   "Delete Charge Code listed"
         Top             =   90
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
         TabIndex        =   15
         Tag             =   "Create a new Charge Code "
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
         TabIndex        =   14
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
         TabIndex        =   9
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
         TabIndex        =   8
         Tag             =   "Edit the information "
         Top             =   105
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
         TabIndex        =   7
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
         TabIndex        =   6
         Tag             =   "Select this Charge Code "
         Top             =   90
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
      Bindings        =   "fxchg.frx":0000
      Height          =   3795
      Left            =   0
      OleObjectBlob   =   "fxchg.frx":0014
      TabIndex        =   0
      Tag             =   "Charge Code Listings"
      Top             =   0
      Width           =   6615
   End
End
Attribute VB_Name = "frmCHARGECODE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNoRecords%
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO



Private Function chkChargeCode()
Dim xCode As String, SQLQ As String, Msg$
Dim rsChargeCode As New ADODB.Recordset

chkChargeCode = False
On Error GoTo chkChargeCode_Err

If Len(txtCode) < 1 Then
    MsgBox lStr("Charge Code Code is a required field")
    txtCode.SetFocus
    Exit Function
End If

If Len(txtDesc) < 1 Then
    MsgBox lStr("Charge Code Description is a required field")
    txtDesc.SetFocus
    Exit Function
End If
If glbLinamar And (Len(txtCode) <> 3 Or Not IsNumeric(txtCode)) Then
    MsgBox lStr("Invalid Charge Code")
    If txtCode.Enabled Then txtCode.SetFocus
    Exit Function
End If
If fglbNewRec% Then
    xCode = CStr(txtCode)
    SQLQ = "SELECT CHARGE_CODE from HR_CHARGE_CODE "
    SQLQ = SQLQ & "WHERE CHARGE_CODE = '" & xCode & "'"
    
    If rsChargeCode.State <> 0 Then rsChargeCode.Close
    rsChargeCode.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If rsChargeCode.BOF And rsChargeCode.EOF Then
        rsChargeCode.Close
    Else
        Msg$ = lStr("This Charge Code number already exists")
        MsgBox Msg$
        rsChargeCode.Close
        Exit Function
    End If
End If

chkChargeCode = True

Exit Function

chkChargeCode_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkChargeCode", "HR_CHARGE_CODE", "Cancel")
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
Dim xCode As String, SQLQ As String, Msg$, a%
Dim rsChargeCode As New ADODB.Recordset

On Error GoTo DelErr

If Len(txtCode) < 1 Then Exit Sub
xCode$ = CStr(txtCode)
If Data1.Recordset.RecordCount = 1 Then
    MsgBox lStr("You can not delete the last Charge Code.")
    Exit Sub
End If


SQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE "
SQLQ = SQLQ & "WHERE AD_CHRGCODE= '" & xCode & "'"

rsChargeCode.Open SQLQ, gdbAdoIhr001, adOpenStatic

If rsChargeCode.BOF And rsChargeCode.EOF Then
    GoTo Lok
Else
    Msg$ = lStr("Employee presently assigned to this Charge Code")
    Msg$ = Msg$ & Chr(10) & ShowEmpnbr(rsChargeCode("ED_EMPNBR"))
'    Msg$ = Msg$ & Chr(10) & rsChargeCode("ED_SURNAME")
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    rsChargeCode.Close
    Exit Sub
End If

Lok:    'looks ok to me
rsChargeCode.Close

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

'Data1.Recordset.Delete
'If Not glbSQL Then Call Pause(0.5)
'Data1.Refresh
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh
Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRPROV", "Delete")
Resume Next

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "CHARGE_CODE = '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
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
txtCode.SetFocus

'Data1.Recordset.AddNew
Call Set_Control("B", Me)
rsDATA.AddNew

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRPROV", "AddNew")
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


Data1.Refresh
Data1.Recordset.Find "CHARGE_CODE='" & xCode & " '"
fglbNewRec% = False

Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPROV", "Update")
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

    RHeading = lStr("Charge Code") & " Listing Report"
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading

    xReport = glbIHRREPORTS & "rgChgCode.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.Formulas(0) = "lblChgCode='" & lStr("Charge Code") & "'"
    Me.vbxCrystal.Connect = RptODBC_SQL

    Me.vbxCrystal.Action = 1


End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()
Dim x
glbCode = Data1.Recordset("CHARGE_CODE")

Unload Me

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_CHARGE_CODE", "SELECT")

End Sub



Private Sub Form_Activate()
Dim xStr

End Sub

Private Sub Form_Load()
Dim SQLQ
glbOnTop = "FRMCHARGECODE"
SQLQ = "SELECT * FROM HR_CHARGE_CODE "
If glbOracle Then
    SQLQ = SQLQ & " ORDER BY UPPER(DESCRIPTION)"
Else
    SQLQ = SQLQ & " ORDER BY DESCRIPTION"
End If

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ
Data1.Refresh
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT
Call modSTUPD(False)
If Not gSec_Upd_Charge_Code Then     'May99 js
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If                          '
'Call setCaption(Me)
'Call setCaption(Me.vbxTrueGrid.Columns.Item(0))
'Call setCaption(Me.vbxTrueGrid.Columns.Item(1))

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

cmdModify.Enabled = FT          'May99 js
cmdFind.Enabled = FT            '
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
                                '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '
                                '
txtCode.Enabled = TF
vbxTrueGrid.Enabled = FT 'Jaddy 11/12/99
txtFindDesc.Enabled = FT        '
txtFindKey.Enabled = FT         '
txtDesc.Enabled = TF            '
                                '
cmdClose.Enabled = FT           '
cmdSelect.Enabled = FT          '
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
'    glbCode = Data1.Recordset("CHARGE_CODE")
    Unload Me
Else
    MsgBox "Save/cancel changes first"
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
        
        SQLQ = "SELECT * FROM HR_CHARGE_CODE "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
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
  
    SQLQ = "select * from HR_CHARGE_CODE WHERE CHARGE_CODE='" & Data1.Recordset!CHARGE_CODE & "'"
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub








