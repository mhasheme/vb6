VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMCounselSet 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Counseling Setup"
   ClientHeight    =   7185
   ClientLeft      =   90
   ClientTop       =   1005
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
   ScaleHeight     =   7185
   ScaleWidth      =   10785
   WindowState     =   2  'Maximized
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "CL_ACTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "00-Comments"
      Top             =   5760
      Width           =   8895
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      DataField       =   "CL_EMAIL2"
      DataSource      =   " "
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
      Index           =   1
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   6
      Tag             =   "00-Email Address"
      Top             =   4800
      Width           =   6780
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      DataField       =   "CL_EMAIL3"
      DataSource      =   " "
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
      Index           =   2
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   7
      Tag             =   "00-Email Address"
      Top             =   5160
      Width           =   6780
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      DataField       =   "CL_EMAIL1"
      DataSource      =   " "
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
      Index           =   0
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   5
      Tag             =   "00-Email Address"
      Top             =   4440
      Width           =   6780
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmcuset.frx":0000
      Height          =   2535
      Left            =   240
      OleObjectBlob   =   "fxmcuset.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   8775
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   9360
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CL_LUSER"
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
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CL_LDATE"
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
      Index           =   0
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3465
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CL_LTIME"
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
      Left            =   7590
      MaxLength       =   8
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9840
      Top             =   5040
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
      GridSource      =   "vbxTrueGrid"
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "CL_REASON"
      Height          =   285
      Index           =   2
      Left            =   1850
      TabIndex        =   4
      Tag             =   "01-Counselling Reason- Code"
      Top             =   4080
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CERE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "CL_TYPE"
      Height          =   285
      Index           =   1
      Left            =   1850
      TabIndex        =   3
      Tag             =   "01-Counselling Type- Code"
      Top             =   3720
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CETY"
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "CL_DIVISION"
      Height          =   285
      Left            =   1850
      TabIndex        =   1
      Tag             =   "00-Specific Division Desired"
      Top             =   3000
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin MSMask.MaskEdBox medTarget 
      DataField       =   "CL_TARGET"
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Tag             =   "10-Target"
      Top             =   3360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Action"
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
      Index           =   7
      Left            =   360
      TabIndex        =   20
      Top             =   5520
      Width           =   840
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address 3"
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
      Index           =   6
      Left            =   360
      TabIndex        =   19
      Top             =   5160
      Width           =   1320
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address 2"
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
      Index           =   5
      Left            =   360
      TabIndex        =   18
      Top             =   4800
      Width           =   1320
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address 1"
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
      Index           =   4
      Left            =   360
      TabIndex        =   17
      Top             =   4440
      Width           =   1440
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   840
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Target"
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
      Left            =   330
      TabIndex        =   14
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Counseling Reason"
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
      Left            =   330
      TabIndex        =   13
      Top             =   4080
      Width           =   1440
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Counseling Type"
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
      Index           =   2
      Left            =   330
      TabIndex        =   12
      Top             =   3720
      Width           =   1635
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "CL_COMPNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   9720
      TabIndex        =   11
      Top             =   4080
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmMCounselSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim DefType(0 To 3)
Dim SystType(0 To 3)
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum



Private Function chkMCounsel()
Dim Msg As String
Dim x%, xchk

chkMCounsel = False

If Len(clpDiv.Text) < 1 Then
    MsgBox lStr("Division must be entered")
    clpDiv.SetFocus
    Exit Function
Else
    If clpDiv.Caption = "Unassigned" And Len(clpDiv.Text) > 0 Then
        MsgBox lStr("Division Code must be valid")
        clpDiv.SetFocus
        Exit Function
    End If
End If

If Len(medTarget) < 1 Then
    MsgBox ("Target must be entered")
    medTarget.SetFocus
    Exit Function
End If
If Not IsNumeric(medTarget) Then
    MsgBox ("Target must be numeric")
    medTarget.SetFocus
    Exit Function
End If

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Counseling Type Code must be entered"
    clpCode(1).SetFocus
    Exit Function
End If

If Len(clpCode(1).Text) > 0 And clpCode(1).Caption = "Unassigned" Then
    MsgBox "Counseling Type Code entered must be known"
    clpCode(1).SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) < 1 Then
    MsgBox "Counseling Reason Code must be entered"
    clpCode(2).SetFocus
    Exit Function
End If

If Len(clpCode(2).Text) > 0 And clpCode(2).Caption = "Unassigned" Then
    MsgBox "Counseling Reason Code entered must be known"
    clpCode(2).SetFocus
    Exit Function
End If

If Len(txtEmail(0)) < 1 Then
    MsgBox ("Email Address 1 must be entered")
    txtEmail(0).SetFocus
    Exit Function
End If


xchk = False


chkMCounsel = True

End Function


Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If


'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
rsDATA.CancelUpdate
Call Display_Value


'Call ST_UPD_MODE(True) ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_COUNSEL_TABLE", "Cancel")
Call RollBack '09June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'Data1.Recordset.Delete
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh


Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_COUNSEL_TABLE", "Delete")
Call RollBack '09June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
clpDiv.SetFocus
Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_COUNSEL_TABLE", "Modify")
Call RollBack '09June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

lblCNum.Caption = "001"

fglbNew = True
Call SET_UP_MODE

clpDiv.SetFocus
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_COUNSEL_TABLE", "Add")
Call RollBack '09June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x%

On Error GoTo cmdOK_Err

If Not chkMCounsel() Then Exit Sub


Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

fglbNew = False
Call SET_UP_MODE

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_COUNSEL_TABLE", "Update")
Call RollBack '09June99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Payroll Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub
Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Payroll Matrix"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpDiv_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub


Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "PAYROLL", "SELECT")

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%

Me.Show
glbOnTop = "frmMCounselSet"

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
If glbFormCaption = "Absence Counseling Setup" Then
    Data1.RecordSource = "SELECT * FROM HR_COUNSEL_ABSENCE ORDER BY CL_DIVISION,CL_TARGET"
Else
    Data1.RecordSource = "SELECT * FROM HR_COUNSEL_LE ORDER BY CL_DIVISION,CL_TARGET"
End If
Data1.Refresh

Call setRptCaption(Me)
Screen.MousePointer = DEFAULT
'Call Display_Value
Call ST_UPD_MODE(False)
                                                '
vbxTrueGrid.Columns(0).Caption = lStr("Division")

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT                           '
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
    MDIMain.panHelp(0).Caption = "Select function from the menu."
Dim I As Integer
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

fUPMode = TF
'vbxTrueGrid.Enabled = FT

clpDiv.Enabled = TF
medTarget.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
txtEmail(0).Enabled = TF
txtEmail(1).Enabled = TF
txtEmail(2).Enabled = TF
memComments.Enabled = TF

End Sub

Private Sub medTarget_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEmail_GotFocus(Index As Integer)
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
        
        If glbFormCaption = "Absence Counseling Setup" Then
            SQLQ = "SELECT * FROM HR_COUNSEL_ABSENCE "
        Else
            SQLQ = "SELECT * FROM HR_COUNSEL_LE "
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err
Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If



Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "Payroll Matrix", "Select")
Call RollBack '09June99 js

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

   
''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    If glbFormCaption = "Absence Counseling Setup" Then
        SQLQ = "SELECT * FROM HR_COUNSEL_ABSENCE where CL_ID= " & Data1.Recordset!CL_ID
    Else
        SQLQ = "SELECT * FROM HR_COUNSEL_LE where CL_ID= " & Data1.Recordset!CL_ID
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
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
UpdateRight = gSec_Matrix
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
ElseIf Data1.Recordset.EOF Then
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
