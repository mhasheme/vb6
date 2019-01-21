VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSManulifeRule 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manulife Transaction Rule Setup"
   ClientHeight    =   7650
   ClientLeft      =   1125
   ClientTop       =   795
   ClientWidth     =   8955
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
   ScaleHeight     =   7650
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFollowPer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "MT_FOLLOWUP_PER"
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
      MaxLength       =   20
      TabIndex        =   28
      Tag             =   "00-Trans Per"
      Top             =   4800
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.TextBox txtFollowRemNum 
      Appearance      =   0  'Flat
      DataField       =   "MT_FOLLOWUP_REMINDER"
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
      Left            =   2595
      MaxLength       =   5
      TabIndex        =   27
      Tag             =   "00-Certificate Number Prefix"
      Top             =   4800
      Width           =   825
   End
   Begin VB.ComboBox cmbFollowPer 
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
      Left            =   4200
      TabIndex        =   26
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txtTranActivity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "MT_TRAN_ACTIVITY"
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
      Left            =   3960
      MaxLength       =   20
      TabIndex        =   25
      Tag             =   "00-Trans Per"
      Top             =   4440
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox cmbTranAct 
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
      Left            =   2600
      TabIndex        =   24
      Top             =   4440
      Width           =   1335
   End
   Begin VB.ComboBox cmbTranPer 
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
      Left            =   4200
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox txtTranRemNum 
      Appearance      =   0  'Flat
      DataField       =   "MT_TRAN_REMINDER"
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
      Left            =   2600
      MaxLength       =   5
      TabIndex        =   2
      Tag             =   "00-Certificate Number Prefix"
      Top             =   4080
      Width           =   825
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6600
      Top             =   7560
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
      TabIndex        =   9
      Top             =   6990
      Width           =   8955
      _Version        =   65536
      _ExtentX        =   15796
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
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   135
         TabIndex        =   10
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Tag             =   "Edit the information "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   12
         Tag             =   "Save changes made"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3540
         TabIndex        =   14
         Tag             =   "Create a new Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4350
         TabIndex        =   15
         Tag             =   "Delete Division listed"
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
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BM_LDATE"
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
      Left            =   5160
      MaxLength       =   25
      TabIndex        =   6
      Text            =   "Ldate"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BM_LTIME"
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
      Left            =   6840
      MaxLength       =   25
      TabIndex        =   7
      Text            =   "LTime"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BM_LUSER"
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
      Left            =   8520
      MaxLength       =   25
      TabIndex        =   8
      Text            =   "LUser"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fsmanurule.frx":0000
      Height          =   2835
      Left            =   120
      OleObjectBlob   =   "fsmanurule.frx":0014
      TabIndex        =   0
      Tag             =   "Division Listings"
      Top             =   0
      Width           =   8745
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "MT_SECTION"
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Tag             =   "01-Benefit - Code"
      Top             =   3000
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BNCD"
      MaxLength       =   10
   End
   Begin VB.TextBox txtTranPer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "MT_TRAN_PER"
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
      MaxLength       =   20
      TabIndex        =   19
      Tag             =   "00-Trans Per"
      Top             =   4080
      Visible         =   0   'False
      Width           =   465
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "MT_EMP"
      Height          =   285
      HelpContextID   =   2
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Tag             =   "00-Enter Status Code"
      Top             =   3360
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "MT_BENEFIT"
      DataSource      =   " "
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   5
      Tag             =   "Benefit Code - Code "
      Top             =   3720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BNCD"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "MT_FOLLOWUP_CODE"
      DataSource      =   " "
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   31
      Tag             =   "01-Follow-up Reason"
      Top             =   5160
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "FURE"
   End
   Begin VB.Label lblFDateRem 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Followup Date Reminder"
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
      TabIndex        =   30
      Top             =   4800
      Width           =   1740
   End
   Begin VB.Label lblFollowPer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Per"
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
      Left            =   3720
      TabIndex        =   29
      Top             =   4800
      Width           =   240
   End
   Begin VB.Label lblTActivity 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Activity"
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
      TabIndex        =   23
      Top             =   4440
      Width           =   1395
   End
   Begin VB.Label lblBen 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit"
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
      TabIndex        =   22
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblEEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   3360
      Width           =   1620
   End
   Begin VB.Label lblTranPer 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Per"
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
      Left            =   3720
      TabIndex        =   20
      Top             =   4080
      Width           =   240
   End
   Begin VB.Label lblTDateRem 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date Reminder"
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
      TabIndex        =   18
      Top             =   4080
      Width           =   2295
   End
   Begin VB.Label lblFollowCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Followup Code"
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
      TabIndex        =   17
      Top             =   5160
      Width           =   1050
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   660
   End
End
Attribute VB_Name = "frmSManulifeRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim xTotRecCount As Long


Private Function chkBenGrp()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset
Dim x
chkBenGrp = False
On Error GoTo chkBenGrp_Err


If Len(clpCode(1).Text) < 1 Then
    MsgBox ("Benefit Group is a required field")
    clpCode(1).SetFocus
    Exit Function
End If
'
'If Len(txtBenAccount.Text) < 1 Then
'    MsgBox ("Benefit Account is a required field")
'    txtBenAccount.SetFocus
'    Exit Function
'End If

'If Len(txtCovClass.Text) < 1 Then
'    MsgBox ("Coverage Class is a required field")
'    txtCovClass.SetFocus
'    Exit Function
'End If


'If fglbNewRec% Then
'    SQLQ = "SELECT * from HR_BENEFITS_GROUP_MATRIX "
'    SQLQ = SQLQ & "WHERE BM_DIV = '" & clpDiv.Text & "'"
'    SQLQ = SQLQ & "AND BM_BENEFIT_GROUP = '" & clpCode(1).Text & "'"
'    SQLQ = SQLQ & "AND BM_BENEFIT_ACCOUNT = '" & txtBenAccount.Text & "'"
'    SQLQ = SQLQ & "AND BM_BENEFIT_CLASS = '" & txtCovClass.Text & "'"
'
'    If snapDivs.State <> 0 Then snapDivs.Close
'    snapDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic
'
'    If snapDivs.BOF And snapDivs.EOF Then
'        snapDivs.Close
'    Else
'        Msg$ = lStr("Duplicate record found!")
'        MsgBox Msg$
'        snapDivs.Close
'        Exit Function
'    End If
'End If

For x = 1 To 1
    If Len(clpCode(x).Text) > 0 And clpCode(x).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(x).SetFocus
        Exit Function
    End If
Next x

chkBenGrp = True

Exit Function

chkBenGrp_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkBenGrp", "HR_Div", "Cancel")
Resume Next

End Function

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub





Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Set_Control("R", Me, rsDATA)


Call modSTUPD(False)  ' reset screen's attributes
cmdClose.SetFocus


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPROv", "Cancel")
Resume Next

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()

glbDiv = ""
glbDivDesc = ""

Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Div As String, SQLQ As String, Msg$, a%

On Error GoTo DelErr


Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


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


Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
'clpCode(1).Enabled = True
clpCode(1).SetFocus
fglbNewRec% = False
'Data1.Recordset.Edit

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '08June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True


Call Set_Control("B", Me)
rsDATA.AddNew


clpCode(1).Enabled = True
clpCode(1).SetFocus


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
Dim xID, ctylist
On Error GoTo OK_Err

If Not chkBenGrp() Then Exit Sub

Call UpdUStats(Me)

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

xID = rsDATA("BM_ID")

Data1.Refresh
Data1.Recordset.Find "BM_ID='" & xID & " '"

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



Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDiv", "SELECT")

End Sub





Private Sub Form_Load()
Dim SQLQ, I, ctylist, x
glbOnTop = "frmSManulifeRule"

'cmbWDate.AddItem ""
'cmbWDate.AddItem "Date of Hire"
'cmbWDate.AddItem "Termination"

cmbTranPer.AddItem "Day"
cmbTranPer.AddItem "Week"
cmbTranPer.AddItem "Month"
cmbTranPer.AddItem "Year"

cmbFollowPer.AddItem "Day"
cmbFollowPer.AddItem "Week"
cmbFollowPer.AddItem "Month"
cmbFollowPer.AddItem "Year"

cmbTranAct.AddItem ""
cmbTranAct.AddItem "Termination"

Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT * FROM HR_MANULIFE_TRAN_RULE "
SQLQ = SQLQ & " ORDER BY MT_SECTION,MT_EMP "
Data1.RecordSource = SQLQ
Data1.Refresh

'Data1.LockType = adLockReadOnly

Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT
Call modSTUPD(False)
If Not gSec_BenefitGroupSetup Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If

'Call setCaption(lblDiv)
'cmdBenDiv.Caption = lStr(cmdBenDiv.Caption)

For I = 1 To 1
    Call setCaption(frmSManulifeRule.vbxTrueGrid.Columns.Item(I))
Next I

Call Display_Value

Call INI_Controls(Me)

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

cmdModify.Enabled = FT
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '


vbxTrueGrid.Enabled = FT
'clpDiv.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
'txtCertiNum.Enabled = TF
cmdClose.Enabled = FT
'cmdPrint.Enabled = FT           '
        
End Sub


Private Sub txtCertiNum_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub vbxTrueGrid_DblClick()
    
If Not Me.vbxTrueGrid.EditActive Then
    glbDiv = Data1.Recordset("DIV")
    glbDivDesc = Data1.Recordset("Division_Name")
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
        
        SQLQ = "select * from HR_DIVISION WHERE " & glbSeleDiv
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


Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'''Sam add July 02 * Remove ADO
Call Display_Value
End Sub
''' Sam add July 2002 * Remove ADO
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        'SQLQ = "select * from HR_DIVISION WHERE DIV='" & Data1.Recordset!Div & "'" & " order by Division_Name"
        SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX "
        SQLQ = SQLQ & "WHERE BM_ID='" & Data1.Recordset!BM_ID & "'"
        SQLQ = SQLQ & " ORDER BY BM_BENEFIT_GROUP,BM_DIV "
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
        Call Set_Control("R", Me, rsDATA)
    End If
    
End Sub

Private Sub IsNew74()
    Dim SQLQ
 
    SQLQ = "select * from HR_DIVISION WHERE DIV='" & Data1.Recordset!Div & "'" & " order by Division_Name"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub

Function GetBonusRptDesc(TablKey)
    Dim rsTABL As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT * FROM WFC_Bonus_Loc_Department WHERE Dept_No = '" & TablKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTABL.EOF And rsTABL.BOF Then
        GetBonusRptDesc = ""
    Else
        GetBonusRptDesc = rsTABL("Dept_Name")
    End If
    rsTABL.Close
End Function

