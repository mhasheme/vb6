VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSAppFormDefaults 
   Caption         =   "Application Form Defaults"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   10335
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmDetails 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7515
      Left            =   240
      TabIndex        =   16
      Top             =   360
      Width           =   9345
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
         DataField       =   "AF_TAB7_LBL"
         Height          =   285
         Index           =   7
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "00-Enter Label for Tab 7"
         Top             =   2880
         Width           =   4995
      End
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
         DataField       =   "AF_TAB6_LBL"
         Height          =   285
         Index           =   6
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "00-Enter Label for Tab 6"
         Top             =   2520
         Width           =   4995
      End
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
         DataField       =   "AF_TAB5_LBL"
         Height          =   285
         Index           =   5
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   8
         Tag             =   "00-Enter Label for Tab 5"
         Top             =   2160
         Width           =   4995
      End
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
         DataField       =   "AF_TAB4_LBL"
         Height          =   285
         Index           =   4
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   6
         Tag             =   "00-Enter Label for Tab 4"
         Top             =   1800
         Width           =   4995
      End
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
         DataField       =   "AF_TAB3_LBL"
         Height          =   285
         Index           =   3
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   4
         Tag             =   "00-Enter Label for Tab 3"
         Top             =   1440
         Width           =   4995
      End
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
         DataField       =   "AF_TAB2_LBL"
         Height          =   285
         Index           =   2
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   2
         Tag             =   "00-Enter Label for Tab 2"
         Top             =   1080
         Width           =   4995
      End
      Begin VB.TextBox txtTab 
         Appearance      =   0  'Flat
         DataField       =   "AF_TAB1_LBL"
         Height          =   285
         Index           =   1
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   0
         Tag             =   "00-Enter Label for Tab 1"
         Top             =   720
         Width           =   4995
      End
      Begin VB.TextBox txtExpDays 
         Appearance      =   0  'Flat
         DataField       =   "AF_ACCT_EXPIRE"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   2680
         MaxLength       =   3
         TabIndex        =   14
         Tag             =   "11-Enter Applicant's Account Expires in Days"
         Top             =   3480
         Width           =   450
      End
      Begin VB.CheckBox chkOpenRequistion 
         Alignment       =   1  'Right Justify
         Caption         =   "Only accept applications for Open Requisitions"
         Height          =   225
         Left            =   570
         TabIndex        =   15
         Tag             =   "40-Check to accept applications for Open Requisitions"
         Top             =   3960
         Width           =   3675
      End
      Begin VB.CheckBox chkHideTab 
         Alignment       =   1  'Right Justify
         Caption         =   "  "
         Height          =   225
         Index           =   6
         Left            =   7800
         TabIndex        =   11
         Tag             =   "40-Check to hide Tab 6"
         Top             =   2550
         Width           =   250
      End
      Begin VB.CheckBox chkHideTab 
         Alignment       =   1  'Right Justify
         Caption         =   "  "
         Height          =   225
         Index           =   5
         Left            =   7800
         TabIndex        =   9
         Tag             =   "40-Check to hide Tab 5"
         Top             =   2190
         Width           =   250
      End
      Begin VB.CheckBox chkHideTab 
         Alignment       =   1  'Right Justify
         Caption         =   "  "
         Height          =   225
         Index           =   7
         Left            =   7800
         TabIndex        =   13
         Tag             =   "40-Check to hide Tab 7"
         Top             =   2910
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.CheckBox chkHideTab 
         Alignment       =   1  'Right Justify
         Caption         =   "  "
         Height          =   225
         Index           =   4
         Left            =   7800
         TabIndex        =   7
         Tag             =   "40-Check to hide Tab 4"
         Top             =   1830
         Width           =   250
      End
      Begin VB.CheckBox chkHideTab 
         Alignment       =   1  'Right Justify
         Caption         =   "  "
         Height          =   225
         Index           =   3
         Left            =   7800
         TabIndex        =   5
         Tag             =   "40-Check to hide Tab 3"
         Top             =   1470
         Width           =   250
      End
      Begin VB.CheckBox chkHideTab 
         Alignment       =   1  'Right Justify
         Caption         =   "  "
         Height          =   225
         Index           =   2
         Left            =   7800
         TabIndex        =   3
         Tag             =   "40-Check to hide Tab 2"
         Top             =   1110
         Width           =   250
      End
      Begin VB.CheckBox chkHideTab 
         Alignment       =   1  'Right Justify
         Caption         =   "  "
         CausesValidation=   0   'False
         Height          =   225
         Index           =   1
         Left            =   7800
         TabIndex        =   1
         Tag             =   "40-Check to hide Tab 1"
         Top             =   750
         Visible         =   0   'False
         Width           =   250
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab 7 change to"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   960
         TabIndex        =   29
         Top             =   2925
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab 6 change to"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   28
         Top             =   2565
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab 5 change to"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   10
         Left            =   960
         TabIndex        =   27
         Top             =   2205
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab 4 change to"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   960
         TabIndex        =   26
         Top             =   1845
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab 3 change to"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   960
         TabIndex        =   25
         Top             =   1485
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab 2 change to"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   960
         TabIndex        =   24
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tab 1 change to"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   960
         TabIndex        =   23
         Top             =   765
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "(999 = Do not expire account)"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   3720
         TabIndex        =   22
         Top             =   3525
         Width           =   2115
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Days"
         Height          =   195
         Index           =   4
         Left            =   3180
         TabIndex        =   21
         Top             =   3525
         Width           =   360
      End
      Begin VB.Label lblTitle 
         Caption         =   "Defaults"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   20
         Top             =   60
         Width           =   2955
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Tab Label Changes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   570
         TabIndex        =   19
         Top             =   360
         Width           =   1665
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Hide"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   7725
         TabIndex        =   18
         Top             =   360
         Width           =   405
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Applicant Account Expires in "
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   17
         Top             =   3525
         Width           =   2070
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   6240
      Top             =   8040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmSAppFormDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean

Sub cmdCancel_Click()
    Dim X
    X = EERetrieve
    Call ST_UPD_MODE(True)
End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdOK_Click()
    Dim X As Integer
    
    If Len(txtExpDays.Text) <> 0 Then
        If Not IsNumeric(txtExpDays.Text) Then
            MsgBox "Applicant Account Expires in Days must be numeric"
            txtExpDays.SetFocus
            Exit Sub
        End If
    End If
    
    If data1.Recordset.EOF Then data1.Recordset.AddNew
    data1.Recordset("AF_ACCT_EXPIRE") = IIf(Not IsNumeric(txtExpDays.Text), 999, txtExpDays.Text)
    data1.Recordset("AF_APPL_OPN_REQ") = IIf(chkOpenRequistion, 1, 0)
    For X = 1 To 7
        data1.Recordset("AF_TAB" & X & "_LBL") = txtTab(X).Text
    Next
    For X = 1 To 7
        data1.Recordset("AF_TAB" & X & "_HIDE") = IIf(chkHideTab(X), 1, 0)
    Next
    data1.Recordset("AF_COMPNO") = "001"
    data1.Recordset("AF_LDATE") = Date
    data1.Recordset("AF_LTIME") = Time$
    data1.Recordset("AF_LUSER") = glbUserID
    data1.Recordset.Update
    
    fglbNew = False
    
    Call SET_UP_MODE
    
    'Call ST_UPD_MODE(False)
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

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
FrmDetails.Enabled = TF
End Sub

Private Sub chkOpenRequistion_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkHideTab_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    Dim X As Integer
    
    glbOnTop = "FRMSAPPFORMDEFAULTS"
    MDIMain.lstPanel.Visible = False
    MDIMain.lstView.Visible = False
    
    data1.ConnectionString = glbAdoIHRDB
    X = EERetrieve
    
    'For x = 1 To chkHideTab.count - 1
    '    chkHideTab(x).Tag = "40-" & chkHideTab(x).Caption
    'Next
    
    'Call ST_UPD_MODE(False)

End Sub

Private Function EERetrieve()
    Dim X As Integer
    
    data1.RecordSource = "HRA_APPLFORM_DFLTS"
    data1.Refresh
    
    If Not data1.Recordset.EOF Then
        txtExpDays.Text = data1.Recordset("AF_ACCT_EXPIRE")
        chkOpenRequistion = IIf(data1.Recordset("AF_APPL_OPN_REQ"), 1, 0)
    Else
        txtExpDays.Text = ""
        chkOpenRequistion = 0
    End If
    
    For X = 1 To 7
        If data1.Recordset.EOF Then
            txtTab(X).Text = ""
        Else
            txtTab(X).Text = data1.Recordset("AF_TAB" & X & "_LBL")
        End If
    Next
    
    For X = 1 To 7
        If data1.Recordset.EOF Then
            chkHideTab(X) = 0
        Else
            chkHideTab(X) = IIf(data1.Recordset("AF_TAB" & X & "_HIDE"), 1, 0)
        End If
    Next
End Function

Public Property Get ChangeAction() As UpdateStateEnum
'If fglbNew Then
'    ChangeAction = NewRecord
'Else
    ChangeAction = OPENING
'End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_AppFormDefaults
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = True
End Property

Public Property Get Deleteble() As Boolean
Deleteble = False
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
    UpdateState = OPENING
    TF = True
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim VR

If Changed = True Then
    VR = MsgBox("Do you want to save changes?", MB_YESNO)
    If VR = IDYES Then
        Me.cmdOK_Click 'Then Pause (0.5) Else isUpdated = False
    ElseIf VR = IDNO Then
        'Call Me.cmdCancel_Click
    End If
End If

Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Function Changed()
Dim rsAppFrmDftl As New ADODB.Recordset
Dim X%, SQLQ

Changed = False

SQLQ = "SELECT * FROM HRA_APPLFORM_DFLTS "
rsAppFrmDftl.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If Not rsAppFrmDftl.EOF Then
    
    'Tab Label change
    For X = 1 To 7
        If UCase(Trim(rsAppFrmDftl("AF_TAB" & X & "_LBL"))) <> UCase(Trim(txtTab(X).Text)) Then
            Changed = True
            Exit For
        End If
    Next
    
    'Hide change
    For X = 2 To 6
        If Abs(Abs(rsAppFrmDftl("AF_TAB" & X & "_HIDE"))) <> chkHideTab(X) Then
            Changed = True
            Exit For
        End If
    Next
    
    'Other
    If Int(rsAppFrmDftl("AF_ACCT_EXPIRE")) <> Int(txtExpDays.Text) Then
        Changed = True
    End If
    
    If Abs(Abs(rsAppFrmDftl("AF_APPL_OPN_REQ"))) <> chkOpenRequistion Then
        Changed = True
    End If
    
End If
rsAppFrmDftl.Close
Set rsAppFrmDftl = Nothing

End Function

Private Sub txtExpDays_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtTab_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub
