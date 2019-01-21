VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUAccrClr 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Clear Accrual File"
   ClientHeight    =   6825
   ClientLeft      =   4395
   ClientTop       =   4080
   ClientWidth     =   11280
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
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6825
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbAccType 
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
      Height          =   315
      ItemData        =   "fuAccrualClr.frx":0000
      Left            =   2800
      List            =   "fuAccrualClr.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "Choose Accrual Type to display"
      Top             =   1800
      Width           =   1935
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   6165
      Width           =   11280
      _Version        =   65536
      _ExtentX        =   19897
      _ExtentY        =   1164
      _StockProps     =   15
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   4800
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   375
         Left            =   7200
         Top             =   120
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
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
   End
   Begin INFOHR_Controls.DateLookup dlpTo 
      Height          =   285
      Left            =   2490
      TabIndex        =   2
      Tag             =   "40-Date upto and including this date forward"
      Top             =   1410
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpFrom 
      Height          =   285
      Left            =   2490
      TabIndex        =   1
      Tag             =   "40-Date from and including this date forward"
      Top             =   1080
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID 
      Height          =   285
      Left            =   2490
      TabIndex        =   0
      Tag             =   "10-Enter Employee Number"
      Top             =   720
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   7195
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2460
      TabIndex        =   4
      Tag             =   "00-Enter Union Code"
      Top             =   2280
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDOR"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
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
      Left            =   600
      TabIndex        =   11
      Top             =   2280
      Width           =   420
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
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
      Left            =   600
      TabIndex        =   10
      Top             =   720
      Width           =   1290
   End
   Begin VB.Label lblFrom 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   1920
      TabIndex        =   9
      Top             =   1125
      Width           =   420
   End
   Begin VB.Label lblTo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   1950
      TabIndex        =   8
      Top             =   1455
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Accrual Type"
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
      Left            =   600
      TabIndex        =   7
      Top             =   1860
      Width           =   945
   End
   Begin VB.Label lblFromTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range"
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
      Left            =   600
      TabIndex        =   6
      Top             =   1125
      Width           =   870
   End
End
Attribute VB_Name = "frmUAccrClr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DeletedRecs As Long

Private Function chkAccrual()
Dim X%
Dim dd&

chkAccrual = False

If Not elpEEID.ListChecker Then
    Exit Function
End If

If Len(dlpFrom.Text) > 0 Then
    If Not IsDate(dlpFrom.Text) Then
        MsgBox "Invalid From Date in Date Range"
        dlpFrom.SetFocus
        Exit Function
    End If
End If

If Len(dlpTo.Text) > 0 Then
    If Not IsDate(dlpTo.Text) Then
        MsgBox "Invalid To Date in Date Range"
        dlpTo.SetFocus
        Exit Function
    End If
End If

If Len(dlpFrom.Text) > 0 And Len(dlpTo.Text) > 0 Then
    dd& = DateDiff("d", CVDate(dlpFrom.Text), CVDate(dlpTo.Text))
    If dd& < 0 Then
        MsgBox "From date must be earlier than To Date"
        dlpFrom.SetFocus
        Exit Function
    End If
End If

If Not clpCode(1).ListChecker Then
    Exit Function
End If

chkAccrual = True
End Function

Public Sub cmdDelete_Click()
Dim a As Integer
Dim SQLQ As String, rc%, DtTm As Variant, X%
Dim DgDef, Title$, Msg$, Response%
Dim recCount As Long

If Not chkAccrual() Then Exit Sub

Title$ = "Mass Accrual File Delete"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.

If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

recCount = getRecordCount_Delete
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Accrual Record " Else Msg$ = Msg$ & " Accrual Records "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Accrual record found to delete."
    Exit Sub
End If

Screen.MousePointer = HOURGLASS

X% = modDelRecs()

Screen.MousePointer = DEFAULT

If DeletedRecs = 0 Then
    MsgBox "No records found for given selection criteria."
Else
    MsgBox DeletedRecs & " records deleted successfully."
End If

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR ACCRUAL", "Delete")
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Function modDelRecs()
Dim SQLQ As String, SQLW As String, iOneWhere As Integer

modDelRecs = False

On Error GoTo cmdDel_Err

Screen.MousePointer = HOURGLASS

SQLQ = "Delete FROM HR_ACCRUAL WHERE 1=1 "

SQLW = ""
If Len(elpEEID.Text) > 0 Then SQLW = SQLW & " AND AC_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
If Len(dlpFrom.Text) > 0 Then SQLW = SQLW & " AND AC_EDATE >= " & Date_SQL(dlpFrom.Text)
If Len(dlpTo.Text) > 0 Then SQLW = SQLW & " AND AC_EDATE <= " & Date_SQL(dlpTo.Text)

If cmbAccType.ListIndex <> 0 Then
    If cmbAccType.Text = "Vacation" Then
        SQLW = SQLW & " AND AC_TYPE= 'VAC'"
    ElseIf cmbAccType.Text = "Sick" Then
        SQLW = SQLW & " AND AC_TYPE= 'SICK'"
    ElseIf cmbAccType.Text = "Compensatory Time" Then
        SQLW = SQLW & " AND AC_TYPE= 'BANK'"
    ElseIf cmbAccType.Text = "Hourly Entitlement" Then
        SQLW = SQLW & " AND AC_TYPE NOT IN ('BANK', 'SICK', 'VAC')"
    End If
End If

If Len(clpCode(1).Text) > 0 Then
    SQLW = SQLW & " AND AC_EMPNBR in (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')) "
End If

SQLQ = SQLQ & SQLW
gdbAdoIhr001.Execute SQLQ, DeletedRecs

modDelRecs = True

Exit Function

cmdDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "modDelRecs", "HR_ACCRUAL", "Delete")

Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If

End Function

Private Sub cmbAccType_Change()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUACCRCLR"
End Sub

Private Sub Form_Load()
Dim SQLQ As String

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

On Error GoTo Ld_Err

glbOnTop = "FRMUACCRCLR"

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT AC_EMPNBR FROM HR_ACCRUAL WHERE AC_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP "
SQLQ = SQLQ & " WHERE " & glbSeleDeptUn & ")"
Data1.RecordSource = SQLQ
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.EOF Then
  MsgBox "ACCRUAL FILE IS EMPTY"
  Unload Me
  Screen.MousePointer = DEFAULT
  Exit Sub
End If

cmbAccType.AddItem "All"
cmbAccType.AddItem "Vacation"
cmbAccType.AddItem "Sick"
cmbAccType.AddItem "Compensatory Time"
cmbAccType.AddItem "Hourly Entitlement"
cmbAccType.ListIndex = 0

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT

Exit Sub

Ld_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Clear Accrual File", "HR_ACCRUAL", "Select")
Resume Next

End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

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

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
Updateble = False
End Property

Public Property Get Deleteble() As Boolean
Deleteble = gSec_Upd_Entitlements
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
Call set_Buttons
MDIMain.MainToolBar.ButtonS(10).Visible = True
MDIMain.MainToolBar.ButtonS(10).Enabled = True

'If Not UpdateRight Then TF = False
'Call ST_UPD_MODE(TF)
End Sub

Private Function getRecordCount_Delete()
    Dim SQLQ As String
    Dim SQLW As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Long
    
    getRecordCount_Delete = 0
    recCount = 0

SQLQ = "SELECT COUNT(AC_ID) AS TOT_REC FROM HR_ACCRUAL WHERE 1=1 "

SQLW = ""
If Len(elpEEID.Text) > 0 Then SQLW = SQLW & " AND AC_EMPNBR in (" & getEmpnbr(elpEEID.Text) & ") "
If Len(dlpFrom.Text) > 0 Then SQLW = SQLW & " AND AC_EDATE >= " & Date_SQL(dlpFrom.Text)
If Len(dlpTo.Text) > 0 Then SQLW = SQLW & " AND AC_EDATE <= " & Date_SQL(dlpTo.Text)

If cmbAccType.ListIndex <> 0 Then
    If cmbAccType.Text = "Vacation" Then
        SQLW = SQLW & " AND AC_TYPE= 'VAC'"
    ElseIf cmbAccType.Text = "Sick" Then
        SQLW = SQLW & " AND AC_TYPE= 'SICK'"
    ElseIf cmbAccType.Text = "Compensatory Time" Then
        SQLW = SQLW & " AND AC_TYPE= 'BANK'"
    ElseIf cmbAccType.Text = "Hourly Entitlement" Then
        SQLW = SQLW & " AND AC_TYPE NOT IN ('BANK', 'SICK', 'VAC')"
    End If
End If

If Len(clpCode(1).Text) > 0 Then
    SQLW = SQLW & " AND AC_EMPNBR in (SELECT ED_EMPNBR FROM HREMP WHERE ED_ORG IN ('" & Replace(clpCode(1).Text, ",", "','") & "')) "
End If

SQLQ = SQLQ & SQLW

    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        recCount = rsEMP("TOT_REC")
    Else
        recCount = 0
    End If
    rsEMP.Close
    Set rsEMP = Nothing
    
    getRecordCount_Delete = recCount

End Function


