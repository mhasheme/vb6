VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNewHire 
   Caption         =   "New Hire Procedure"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   10335
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmDetails 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7515
      Left            =   240
      TabIndex        =   32
      Top             =   360
      Width           =   9345
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "User Defined Table"
         Height          =   225
         Index           =   35
         Left            =   840
         TabIndex        =   15
         Top             =   6000
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Employee ADP Data"
         Height          =   225
         Index           =   8
         Left            =   840
         TabIndex        =   41
         Top             =   3240
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Succession Planning"
         Height          =   225
         Index           =   14
         Left            =   840
         TabIndex        =   40
         Top             =   5640
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Languages"
         Height          =   225
         Index           =   13
         Left            =   840
         TabIndex        =   39
         Top             =   5310
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Employee Flags"
         Height          =   225
         Index           =   5
         Left            =   840
         TabIndex        =   5
         Top             =   2160
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "G/L Distribution"
         Height          =   225
         Index           =   4
         Left            =   840
         TabIndex        =   4
         Top             =   1800
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Other Information"
         Height          =   225
         Index           =   7
         Left            =   840
         TabIndex        =   7
         Top             =   2880
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Affirmative Action"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   34
         Left            =   4110
         TabIndex        =   38
         Top             =   6870
         Visible         =   0   'False
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Benefits/Beneficiary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   240
         TabIndex        =   12
         Top             =   6300
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Dollar Entitlement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   16
         Left            =   240
         TabIndex        =   13
         Top             =   6600
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Other Earnings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   17
         Left            =   240
         TabIndex        =   14
         Top             =   6930
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Position  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   18
         Left            =   4140
         TabIndex        =   16
         Top             =   360
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Performance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   20
         Left            =   4140
         TabIndex        =   18
         Top             =   1050
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Salary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   19
         Left            =   4140
         TabIndex        =   17
         Top             =   720
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Attendance"
         Height          =   225
         Index           =   21
         Left            =   4740
         TabIndex        =   19
         Top             =   1740
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Vacation and Sick Entitlements"
         Height          =   225
         Index           =   22
         Left            =   4740
         TabIndex        =   20
         Top             =   2100
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Associations"
         Height          =   225
         Index           =   12
         Left            =   840
         TabIndex        =   11
         Top             =   4980
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Continuing Education"
         Height          =   225
         Index           =   11
         Left            =   840
         TabIndex        =   10
         Top             =   4650
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Formal Education"
         Height          =   225
         Index           =   10
         Left            =   840
         TabIndex        =   9
         Top             =   4320
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Skills"
         Height          =   225
         Index           =   9
         Left            =   840
         TabIndex        =   8
         Top             =   3960
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Payroll/Banking"
         Height          =   225
         Index           =   6
         Left            =   840
         TabIndex        =   6
         Top             =   2520
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Dependents"
         Height          =   225
         Index           =   3
         Left            =   840
         TabIndex        =   3
         Top             =   1440
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Contacts"
         Height          =   225
         Index           =   2
         Left            =   840
         TabIndex        =   2
         Top             =   1080
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Status/Dates"
         Height          =   225
         Index           =   1
         Left            =   840
         TabIndex        =   1
         Top             =   750
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Demographics"
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   840
         TabIndex        =   0
         Top             =   390
         Value           =   1  'Checked
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Hourly Entitlements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   23
         Left            =   4140
         TabIndex        =   21
         Top             =   2460
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Claims/Medical Information"
         Height          =   225
         Index           =   28
         Left            =   4740
         TabIndex        =   26
         Top             =   4620
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Corrective Action"
         Height          =   225
         Index           =   27
         Left            =   4740
         TabIndex        =   25
         Top             =   4260
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Root Cause"
         Height          =   225
         Index           =   26
         Left            =   4740
         TabIndex        =   24
         Top             =   3900
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Injury/Location"
         Height          =   225
         Index           =   25
         Left            =   4740
         TabIndex        =   23
         Top             =   3540
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Incident Data"
         Height          =   225
         Index           =   24
         Left            =   4740
         TabIndex        =   22
         Top             =   3180
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Employee Follow-up"
         Height          =   225
         Index           =   32
         Left            =   4710
         TabIndex        =   30
         Top             =   6240
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "WSIB Cost Information"
         Height          =   225
         Index           =   30
         Left            =   4740
         TabIndex        =   28
         Top             =   5280
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Contact Information"
         Height          =   225
         Index           =   29
         Left            =   4740
         TabIndex        =   27
         Top             =   4950
         Width           =   2600
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Comments"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   31
         Left            =   4140
         TabIndex        =   29
         Top             =   5610
         Width           =   3200
      End
      Begin VB.CheckBox chkItem 
         Alignment       =   1  'Right Justify
         Caption         =   "Cobra Maintenance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   33
         Left            =   4110
         TabIndex        =   31
         Top             =   6540
         Visible         =   0   'False
         Width           =   3200
      End
      Begin VB.Label lblTitle 
         Caption         =   " Follow-up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   4140
         TabIndex        =   37
         Top             =   5940
         Width           =   2955
      End
      Begin VB.Label lblTitle 
         Caption         =   " Employee Basic Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   36
         Top             =   60
         Width           =   2955
      End
      Begin VB.Label lblTitle 
         Caption         =   " Education and Skills"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   35
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label lblTitle 
         Caption         =   " Attendance and Entitlements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4125
         TabIndex        =   34
         Top             =   1410
         Width           =   2955
      End
      Begin VB.Label lblTitle 
         Caption         =   " Health and Safety"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4125
         TabIndex        =   33
         Top             =   2820
         Width           =   2955
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
Attribute VB_Name = "frmNewHire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ChangeCBox

Private Sub chkItem_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub chkItem_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Index = 0 Then chkItem(0) = 0
End Sub

Private Sub chkItem_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
If Index = 0 Then chkItem(0) = 1
End Sub

Sub cmdCancel_Click()
Dim x
x = EERetrieve
Call ST_UPD_MODE(True)
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


'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
With data1.Recordset
    .MoveFirst
    For x = 1 To chkItem.count '- 1
        .MoveFirst
        .Find "ID=" & x & ""
        If Not .EOF Then
            !NewHire = IIf(chkItem(x - 1), 1, 0)
            .Update
        End If
    Next
End With
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

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Click()
'Dim x
'For x = 0 To chkItem.count - 1
'    With Data1.Recordset
'        .AddNew
'        !ID = x + 1
'        !NewHire = IIf(chkItem(x), 1, 0)
'        !MenuItem = chkItem(x).Caption
'        .Update
'    End With
'Next
End Sub

Private Sub Form_Load()
glbOnTop = "FRMNEWHIRE"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim x
If glbCountry = "U.S.A." Then
    chkItem(32).Visible = True
    ' danielk - 12/31/2002 - Added EEO
    chkItem(33).Visible = True
End If
data1.ConnectionString = glbAdoIHRDB
x = EERetrieve
For x = 0 To chkItem.count - 1
    chkItem(x).Tag = "40-" & chkItem(x).Caption
Next

'Call ST_UPD_MODE(False)
If Not gSec_Upd_New_Hire Then
 '   cmdModify.Enabled = False
End If

chkItem(4).Caption = lStr(chkItem(4).Caption)
chkItem(26).Caption = lStr(chkItem(26).Caption)
chkItem(30).Caption = lStr(chkItem(30).Caption)
chkItem(35).Caption = lStr("User Defined Table") 'Ticket #30482 Franks 08/15/2017

End Sub

Private Function EERetrieve()
Dim dblbl

data1.RecordSource = "HRNEWHIRE"
data1.Refresh
For x = 1 To chkItem.count '- 1
    With data1.Recordset
        .MoveFirst

        .Find "ID=" & x & ""
        
        If .EOF Then
             chkItem(x - 1) = 0
        Else
            chkItem(x - 1) = IIf(!NewHire, 1, 0)
        End If
    End With
Next
End Function


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
UpdateRight = gSec_Upd_New_Hire
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
Call ChkCBoxChange
If ChangeCBox = True Then
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

Private Sub ChkCBoxChange()
Dim rsCB As New ADODB.Recordset
Dim x%, SQLQ

SQLQ = "SELECT ID,MENUITEM, NEWHIRE "
SQLQ = SQLQ & " FROM HRNEWHIRE "
rsCB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
ChangeCBox = False
Do Until rsCB.EOF
    If Abs(Abs(rsCB("NEWHIRE"))) <> chkItem(rsCB("ID") - 1) Then
        GoTo theend
    End If
    rsCB.MoveNext
Loop
Exit Sub
theend:
ChangeCBox = True
End Sub

