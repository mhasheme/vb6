VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPosEval 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Evaluation Factors for Position"
   ClientHeight    =   6300
   ClientLeft      =   105
   ClientTop       =   -1515
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
   NegotiateMenus  =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   6300
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "JE_FACTOR"
      Height          =   285
      Index           =   1
      Left            =   1725
      TabIndex        =   1
      Tag             =   "01-Evaluation Factor - Code"
      Top             =   2460
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBFC"
   End
   Begin VB.TextBox medPoints 
      Appearance      =   0  'Flat
      DataField       =   "JE_POINTS"
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
      Left            =   6300
      MaxLength       =   9
      TabIndex        =   5
      Tag             =   "11-Points"
      Top             =   3090
      Width           =   2055
   End
   Begin VB.TextBox medLevel 
      Appearance      =   0  'Flat
      DataField       =   "JE_SUBLVL"
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
      Left            =   6300
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "00-Sub_Level"
      Top             =   2760
      Width           =   645
   End
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataField       =   "JE_COMMENTS"
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
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   6
      Tag             =   "00-Comments on this Factor"
      Top             =   3450
      Width           =   3600
   End
   Begin VB.TextBox txtWeight 
      Appearance      =   0  'Flat
      DataField       =   "JE_WGT"
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
      Left            =   2040
      TabIndex        =   4
      Tag             =   "10-Weight Factor Carries"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox medLevel1 
      Appearance      =   0  'Flat
      CausesValidation=   0   'False
      DataField       =   "JE_LEVEL"
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
      Left            =   2040
      MaxLength       =   2
      TabIndex        =   2
      Tag             =   "00-Level"
      Top             =   2790
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JE_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1470
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5565
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JE_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3150
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5565
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JE_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   4200
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5565
      Visible         =   0   'False
      Width           =   900
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fspeval.frx":0000
      Height          =   1695
      Left            =   60
      OleObjectBlob   =   "Fspeval.frx":0014
      TabIndex        =   0
      Tag             =   "Evaluation Factors Lookup"
      Top             =   600
      Width           =   9120
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblPosDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Descr"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2880
         TabIndex        =   22
         Top             =   120
         Width           =   630
      End
      Begin VB.Label lblPosition 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ABCD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         TabIndex        =   21
         Top             =   135
         Width           =   630
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   165
         Width           =   690
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   6480
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportSource    =   1
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   5760
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
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Points"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   4995
      TabIndex        =   18
      Top             =   3060
      Width           =   540
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sub-Level"
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
      Left            =   4995
      TabIndex        =   17
      Top             =   2775
      Width           =   870
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Left            =   330
      TabIndex        =   16
      Top             =   3465
      Width           =   870
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Left            =   330
      TabIndex        =   15
      Top             =   3150
      Width           =   615
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
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
      TabIndex        =   14
      Top             =   2835
      Width           =   390
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Factor"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   13
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label lblID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblPositions 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "POST"
      DataField       =   "JE_CODE"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   750
      TabIndex        =   11
      Top             =   5580
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "JE_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   5580
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "frmPosEval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fglbRecords%, fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Private Function chkPosEval()
Dim SQLQ As String, Msg As String, dd#, PID&, Factor$

chkPosEval = False

On Error GoTo chkPosEval_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Evaluation Factor is a required field"
     clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Evaluation Factor must be valid"
     clpCode(1).SetFocus
    Exit Function
End If

If (Not IsNumeric(txtWeight)) And Len(Trim(txtWeight)) > 0 Then
    MsgBox "Weight should be numeric"
    txtWeight.SetFocus
    Exit Function
End If

If Len(Trim(medPoints)) = 0 Then
    MsgBox "Points is a required field"
    medPoints.SetFocus
    Exit Function
End If

If Not IsNumeric(medPoints) Then
    MsgBox "Points should be numeric"
    medPoints.SetFocus
    Exit Function
End If

PID& = CLng(Val(lblID))
Factor$ = clpCode(1).Text

If modISDupFactor(glbPos$, Factor$, PID&) Then
    MsgBox "Factor must be unique"
     clpCode(1).SetFocus
    Exit Function
End If

chkPosEval = True

Exit Function

chkPosEval_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HRJOBEVL", "edit/Add")
Call RollBack '15June99 js

End Function

Public Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control
rsDATA.CancelUpdate
Call Display_Value



'Call ST_UPD_MODE(True)  ' reset screen's attributes


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack '15June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdClose_Click()
glbUserUploadMode = SwitchForm: Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    fglbRecords% = False
    Exit Sub
Else
    fglbRecords% = True
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRJOBEVL", "Delete")
Call RollBack '15June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdModify_Click()

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If

fglbEditMode% = True

On Error GoTo Mod_Err

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
'clpCode(1).Enabled = True
'clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '15June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_Job_Master Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If
fglbNew = True
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

On Error GoTo AddN_Err


'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)
rsDATA.AddNew


fglbEditMode% = True
lblCNum.Caption = "001"
lblPositions.Caption = glbPos$

clpCode(1).SetFocus
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRJOBEVL", "Add")
Call RollBack '15June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdOK_Click()
Dim xChange

On Error GoTo OK_Err
fglbNew = False
If Not chkPosEval() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

fglbEditMode% = False

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBEVL", "Update")
Call RollBack '15June99 js
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

'RHeading = Me.Caption & lblPosDesc.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub
Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
RHeading = Mid(RHeading, 1, InStr(RHeading, "-"))
RHeading = RHeading & " " & lblPosDesc.Caption

'RHeading = Me.Caption & lblPosDesc.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Form_Activate()
Call ST_UPD_MODE(True)
If fglbRecords Then Data1.Recordset.MoveFirst
Call SET_UP_MODE
Screen.MousePointer = DEFAULT

glbOnTop = "FRMPOSEVAL"
End Sub

Private Sub Form_Load()

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%

glbOnTop = "FRMPOSEVAL"

Screen.MousePointer = HOURGLASS

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False


Data1.ConnectionString = glbAdoIHRDB
If glbWFC Then 'Ticket #25911 Franks 10/21/2014
    If glbPos = "" Then frmJOBSWFC.Show 1
Else
    If glbPos = "" Then frmJOBS.Show 1
End If
If glbPos = "" Then glbUserUploadMode = UploadFormWithoutCheck: Unload Me: Exit Sub

lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$
Me.Caption = "Evaluation Factors Position - " & lblPosition


Screen.MousePointer = DEFAULT

X% = EERetrieve()     '(glbPos$)  ' sets fglbRecords
Call Display_Value
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

    MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
End Sub

Private Sub medLevel_GotFocus(Index As Integer)

Call SetPanHelp(ActiveControl)

End Sub

Private Sub medLevel1_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPoints_GotFocus()
    Call SetPanHelp(ActiveControl)
    MDIMain.panHelp(2).Caption = " "
End Sub

Private Sub medPoints_LostFocus()
   
If Not IsNull(medPoints) Then
    If IsNumeric(medPoints) And medPoints <> "0" Then
        medPoints = Round(medPoints, 2)
    End If
End If

End Sub

Private Sub mnu_File_Close_Click()
    Unload Me
End Sub

Private Sub mnu_File_Exit_Click()
    Call ApplicationEnd
End Sub

Private Sub mnu_F_PrintSetup_Click()
    MDIMain.vbxCommonDlg.Action = 5
End Sub

Private Sub mnu_Skills_Click()

Load frmPosSkills
frmPosSkills.ZOrder BRINGTOFRONT

End Sub

Private Sub mnu_WIN_about_Click()
'    MenuAbout
frmAbout.Show 1
End Sub

Private Sub mnu_WIN_Arrange_Click()
    MDIMain.Arrange 3
End Sub

Private Sub mnu_WIN_Cascade_Click()
    MDIMain.Arrange 0
End Sub

Private Sub mnu_WIN_TILEH_Click()
    MDIMain.Arrange 1
End Sub

Private Sub mnu_WIN_TILEV_Click()
    MDIMain.Arrange 2
End Sub

Public Function EERetrieve()    '(StrPos$)
Dim SQLQ$

EERetrieve = False
Screen.MousePointer = HOURGLASS

On Error GoTo EERetrieveErr


' out or left join query not updateable - so do straight.
SQLQ$ = "SELECT * FROM HRJOBEVL "
SQLQ$ = SQLQ$ & "WHERE JE_CODE = '" & glbPos$ & "'"       'StrPos$ & "'"
Data1.RecordSource = SQLQ$
Data1.Refresh
lblPosition.Caption = glbPos$
lblPosDesc.Caption = glbPosDesc$


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERetrieveErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Pos Skills", "HRJOBSK", "SELECT")
Call RollBack '15June99 js

Exit Function

End Function

Private Function modISDupFactor(Pos$, Factor$, PID&)
Dim SQLQ$
Dim snapEval As New ADODB.Recordset

modISDupFactor = True

On Error GoTo modISDupFactor_Err

Screen.MousePointer = HOURGLASS

SQLQ$ = "SELECT * FROM HRJOBEVL "
SQLQ$ = SQLQ$ & "Where (JE_CODE = '" & Pos$ & "' "
SQLQ$ = SQLQ$ & "AND JE_FACTOR = '" & Factor$ & "' "
SQLQ$ = SQLQ$ & "AND JE_ID <> " & PID& & ") "


snapEval.Open SQLQ$, gdbAdoIhr001, adOpenStatic

If snapEval.BOF And snapEval.EOF Then
    modISDupFactor = False
End If

snapEval.Close
Screen.MousePointer = DEFAULT
Exit Function

modISDupFactor_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Code Snap", "TABL", "SELECT")
Call RollBack '15June99 js

End Function



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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdPrint.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdModify.Enabled = FT

'vbxTrueGrid.Enabled = FT
 clpCode(1).Enabled = TF
medLevel1(0).Enabled = TF
txtWeight.Enabled = TF
txtComments.Enabled = TF
medLevel(1).Enabled = TF
medPoints.Enabled = TF
'If Data1.Recordset.EOF Or Data1.Recordset.EOF Then
'    cmdDelete.Enabled = False
'    cmdModify.Enabled = False
'End If
End Sub



Private Sub txtComments_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtWeight_GotFocus()

Call SetPanHelp(ActiveControl)
MDIMain.panHelp(2).Caption = " "    'laura jan 06, 1997

End Sub


Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
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
        
        ' out or left join query not updateable - so do straight.
        SQLQ$ = "SELECT * FROM HRJOBEVL "
        SQLQ$ = SQLQ$ & "WHERE JE_CODE = '" & glbPos$ & "'"       'StrPos$ & "'"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
    '    cmdOK.SetFocus
'    Else
    '    cmdClose.SetFocus
'    End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, X%
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value

 ' set description for code

If Data1.Recordset.RecordCount <> 0 Then
    If Not IsNull(Data1.Recordset("JE_WGT")) Then
        txtWeight = Data1.Recordset("JE_WGT")
    Else
        txtWeight = ""
    End If
End If
    
Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRJOBSKL", "Add")
Call RollBack '15June99 js

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT
Call Display_Value

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

''' Sam add July 2002 * Remove Binding Control
Public Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Me.cmdModify_Click
        Exit Sub
    End If
    
    
    SQLQ = "SELECT * FROM HRJOBEVL "
    SQLQ = SQLQ & "WHERE JE_ID = " & Data1.Recordset!JE_ID

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    lblID = rsDATA!JE_ID

    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
    Me.cmdModify_Click
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property
Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelatePOS
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Job_Master
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
clpCode(1).Enabled = TF
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub

Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

