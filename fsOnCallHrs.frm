VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSOnCallHrs 
   Caption         =   "On Call Hours Setup"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   9870
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3900
      Top             =   7770
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
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
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "CH_HRS1"
      DataSource      =   "data1"
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   7
      Tag             =   "11-Hours for Sunday"
      Top             =   960
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "CH_HRS2"
      DataSource      =   "data1"
      Height          =   285
      Index           =   2
      Left            =   3720
      TabIndex        =   8
      Tag             =   "11-Hours for Monday"
      Top             =   1335
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "CH_HRS3"
      DataSource      =   "data1"
      Height          =   285
      Index           =   3
      Left            =   3720
      TabIndex        =   9
      Tag             =   "11-Hours for Tuesday"
      Top             =   1725
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "CH_HRS4"
      DataSource      =   "data1"
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   10
      Tag             =   "11-Hours for Wednesday"
      Top             =   2100
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "CH_HRS5"
      DataSource      =   "data1"
      Height          =   285
      Index           =   5
      Left            =   3720
      TabIndex        =   11
      Tag             =   "11-Hours for Thursday"
      Top             =   2475
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "CH_HRS6"
      DataSource      =   "data1"
      Height          =   285
      Index           =   6
      Left            =   3720
      TabIndex        =   12
      Tag             =   "11-Hours for Friday"
      Top             =   2865
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "CH_HRS7"
      DataSource      =   "data1"
      Height          =   285
      Index           =   7
      Left            =   3720
      TabIndex        =   13
      Tag             =   "11-Hours for Saturday"
      Top             =   3240
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOCHours 
      DataField       =   "CH_OCHRS1"
      DataSource      =   "data1"
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   0
      Tag             =   "11-Hours for Sunday"
      Top             =   960
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOCHours 
      DataField       =   "CH_OCHRS2"
      DataSource      =   "data1"
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   1
      Tag             =   "11-Hours for Monday"
      Top             =   1340
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOCHours 
      DataField       =   "CH_OCHRS3"
      DataSource      =   "data1"
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   2
      Tag             =   "11-Hours for Tuesday"
      Top             =   1720
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOCHours 
      DataField       =   "CH_OCHRS4"
      DataSource      =   "data1"
      Height          =   285
      Index           =   4
      Left            =   2400
      TabIndex        =   3
      Tag             =   "11-Hours for Wednesday"
      Top             =   2100
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOCHours 
      DataField       =   "CH_OCHRS5"
      DataSource      =   "data1"
      Height          =   285
      Index           =   5
      Left            =   2400
      TabIndex        =   4
      Tag             =   "11-Hours for Thursday"
      Top             =   2480
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOCHours 
      DataField       =   "CH_OCHRS6"
      DataSource      =   "data1"
      Height          =   285
      Index           =   6
      Left            =   2400
      TabIndex        =   5
      Tag             =   "11-Hours for Friday"
      Top             =   2860
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medOCHours 
      DataField       =   "CH_OCHRS7"
      DataSource      =   "data1"
      Height          =   285
      Index           =   7
      Left            =   2400
      TabIndex        =   6
      Tag             =   "11-Hours for Saturday"
      Top             =   3240
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "# of Hours On Call"
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
      Height          =   420
      Left            =   2400
      TabIndex        =   23
      Top             =   500
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Hours to Pay"
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
      Height          =   195
      Left            =   3720
      TabIndex        =   22
      Top             =   675
      Width           =   1110
   End
   Begin VB.Label lblWeek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sunday"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   21
      Top             =   1005
      Width           =   540
   End
   Begin VB.Label lblWeek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Monday"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1020
      TabIndex        =   20
      Top             =   1380
      Width           =   570
   End
   Begin VB.Label lblWeek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tuesday"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1020
      TabIndex        =   19
      Top             =   1770
      Width           =   615
   End
   Begin VB.Label lblWeek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Wednesday"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   1020
      TabIndex        =   18
      Top             =   2145
      Width           =   855
   End
   Begin VB.Label lblWeek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Thursday"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   1020
      TabIndex        =   17
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label lblWeek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Friday"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   1020
      TabIndex        =   16
      Top             =   2910
      Width           =   420
   End
   Begin VB.Label lblWeek 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Saturday"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   1020
      TabIndex        =   15
      Top             =   3285
      Width           =   630
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Statutory Holidays will replace the Hours to Pay with 4."
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
      Height          =   195
      Left            =   1020
      TabIndex        =   14
      Top             =   3840
      Width           =   5220
   End
End
Attribute VB_Name = "frmSOnCallHrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean

Sub cmdCancel_Click()
Dim x
x = EERetrieve
Call ST_UPD_MODE(True)
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Sub cmdModify_Click()
Call ST_UPD_MODE(True)
End Sub


Sub cmdOK_Click()

Dim x As Integer

'data1.RecordSource = "SELECT * FROM HR_ONCALL_HOURS"
'data1.Refresh

    If Not data1.Recordset.EOF Then
        data1.Recordset.MoveFirst
    Else
        'data1.Recordset.Find "CH_ID=" & data1.Recordset("CH_ID")
        data1.Recordset.AddNew
    End If
        
    data1.Recordset("CH_COMPNO") = "001"
    For x = 1 To medHours.count
        If medHours(x) = "" Then
            data1.Recordset("CH_HRS" & x) = 0
        Else
            data1.Recordset("CH_HRS" & x) = medHours(x)
        End If
    Next
    For x = 1 To medOCHours.count
        If medOCHours(x) = "" Then
            data1.Recordset("CH_OCHRS" & x) = 0
        Else
            data1.Recordset("CH_OCHRS" & x) = medOCHours(x)
        End If
    Next
    data1.Recordset("CH_LDATE") = Date
    data1.Recordset("CH_LTIME") = Time$
    data1.Recordset("CH_LUSER") = glbUserID
    data1.Recordset.Update

data1.Refresh

fglbNew = False
Call SET_UP_MODE

End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer
Dim FT As Integer
Dim x As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

For x = 1 To medHours.count
    medHours(x).Enabled = TF
Next
For x = 1 To medOCHours.count
    medOCHours(x).Enabled = TF
Next

End Sub

Private Sub Form_Activate()
fglbNew = False
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = Me.name

Dim x

data1.ConnectionString = glbAdoIHRDB

x = EERetrieve

fglbNew = False

Call SET_UP_MODE

End Sub

Private Function EERetrieve()
Dim rsOnCall As New ADODB.Recordset
Dim x As Integer

On Error Resume Next

rsOnCall.Open "HR_ONCALL_HOURS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsOnCall.EOF Then
    rsOnCall.AddNew
    rsOnCall("CH_COMPNO") = "001"
    
    For x = 1 To medHours.count
    '    If medHours(x) = "" Then
            rsOnCall("CH_HRS" & x) = 0
    '    Else
    '        rsOnCall("CH_HRS" & x) = medHours(x)
    '    End If
    Next
    For x = 1 To medOCHours.count
    '    If medOCHours(x) = "" Then
            rsOnCall("CH_OCHRS" & x) = 0
    '    Else
    '        rsOnCall("CH_HRS" & x) = medOCHours(x)
    '    End If
    Next
    rsOnCall("CH_LDATE") = Date
    rsOnCall("CH_LTIME") = Time$
    rsOnCall("CH_LUSER") = glbUserID
    rsOnCall.Update
End If

data1.RecordSource = "SELECT * FROM HR_ONCALL_HOURS"
data1.Refresh

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub medHours_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub medHours_KeyPress(Index As Integer, KeyAscii As Integer)
'    If Chr(KeyAscii) = "'" Or Chr(KeyAscii) = """" Then KeyAscii = 0
'End Sub

Private Sub medHours_LostFocus(Index As Integer)
    If Trim(medHours(Index)) = "" Then medHours(Index) = 0
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
UpdateRight = gSec_Upd_OnCallHours
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

If fglbNew Then
    UpdateState = NewRecord
    TF = True
Else
    UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

