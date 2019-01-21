VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUSB 
   Caption         =   "Union Sick Bank"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   10845
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      DataField       =   "WU_EFDATE"
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   10
      Tag             =   "40- Date Range used to select records"
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "WU_ORG"
      Height          =   285
      Index           =   0
      Left            =   2520
      TabIndex        =   9
      Tag             =   "01-Union Code"
      Top             =   2880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "WU_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5160
      MaxLength       =   25
      TabIndex        =   4
      Text            =   "LUser"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "WU_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3480
      MaxLength       =   25
      TabIndex        =   3
      Text            =   "LTime"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "WU_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1800
      MaxLength       =   25
      TabIndex        =   2
      Text            =   "Ldate"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   7260
      Width           =   10845
      _Version        =   65536
      _ExtentX        =   19129
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
      Begin VB.CommandButton cmdRecalc 
         Caption         =   "Recalculate"
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
         Left            =   360
         TabIndex        =   1
         Top             =   90
         Width           =   1575
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9000
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7440
      Top             =   6720
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmUSB.frx":0000
      Height          =   2595
      Left            =   120
      OleObjectBlob   =   "frmUSB.frx":0014
      TabIndex        =   5
      Tag             =   "Division Listings"
      Top             =   120
      Width           =   7935
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "WU_USB"
      Height          =   285
      Left            =   2830
      TabIndex        =   12
      Tag             =   "11-Hours for this reason "
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
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
   Begin INFOHR_Controls.DateLookup dlpDateRange 
      DataField       =   "WU_ETDATE"
      Height          =   285
      Index           =   1
      Left            =   2520
      TabIndex        =   11
      Tag             =   "40- Date Range used to select records"
      Top             =   3600
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label valOuts 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2830
      TabIndex        =   17
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label valTaken 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Taken"
      DataField       =   "WU_USBT"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2830
      TabIndex        =   16
      Top             =   4350
      Width           =   465
   End
   Begin VB.Label lblOuts 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   15
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label lblTaken 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Taken"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   4350
      Width           =   465
   End
   Begin VB.Label lblBank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   3960
      Width           =   900
   End
   Begin VB.Label lblFromDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sick Entitilement:   From Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   3240
      Width           =   2085
   End
   Begin VB.Label lblToDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1680
      TabIndex        =   7
      Top             =   3600
      Width           =   585
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   2910
      Width           =   660
   End
End
Attribute VB_Name = "frmUSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Code_Snap(1) As New ADODB.Recordset
Dim CodeCodes(1, 2)
Dim SFDATE, STDATE
Dim rsDATA As New ADODB.Recordset
Dim fglbNew As Boolean

Public Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

fglbNew = False

rsDATA.CancelUpdate
Call Display_Value
 
Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "WHSCC_USB", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSelect_Click()
Unload Me
End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, X


If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

fglbNew = False

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

Call ST_UPD_MODE(False)

Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "WHSCC_USB", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Public Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
'clpCode(0).SetFocus
Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "WHSCC_USB", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String


Call ST_UPD_MODE(True)

On Error GoTo AddN_Err

fglbNew = True

Call Set_Control("B", Me)
rsDATA.AddNew

dlpDateRange(0).Text = SFDATE
dlpDateRange(1).Text = STDATE
valOuts.Caption = ""


Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OHS_ROOT_CAUSE", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Public Sub cmdOK_Click()
Dim bmk As Variant
Dim xID As Long
On Error GoTo Add_Err

If Not chkWHSCC() Then Exit Sub
Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

Call ReCalcUSB("", "ORG")
xID = rsDATA("WU_ID") 'Data1.Recordset("WU_ID")
Data1.Refresh

Data1.Recordset.Find "WU_ID=" & xID

fglbNew = False

Call ST_UPD_MODE(False)

Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "WHSCC_USB", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub


Public Sub cmdPrint_Click()
Dim RHeading As String
Me.vbxCrystal.Destination = crptToPrinter
RHeading = "Union Sick Bank"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Action = 1
End Sub
Public Sub cmdView_Click()
Dim RHeading As String
Me.vbxCrystal.Destination = crptToWindow
RHeading = "Union Sick Bank"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Action = 1
End Sub

Private Sub cmdRecalc_Click()
Call ReCalcUSB("", "ORG")
Data1.Refresh
vbxTrueGrid.SetFocus
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%
Dim xtime, SQLQ As String, numInc%
Dim rsPARCO As New ADODB.Recordset
On Error GoTo Err_Deal
glbOnTop = "FRMUSB"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Screen.MousePointer = HOURGLASS
Data1.ConnectionString = glbAdoIHRDB

Call setCaption(lblTitle(0))

Call EEList

Call ST_UPD_MODE(False)

rsPARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
If Not rsPARCO.EOF Then
    SFDATE = rsPARCO("PC_FDATES")
    STDATE = rsPARCO("PC_TDATES")
End If
rsPARCO.Close

If Not gSec_Upd_WHSCC_USB Then
    cmdRecalc.Enabled = False
End If
    
Call INI_Controls(Me)

Screen.MousePointer = DEFAULT
Exit Sub

Err_Deal:
Debug.Print "Test"
End Sub

Private Function EEList()
Dim SQLQ As String, Q As QueryDef
Dim countr   As Integer  ' EEList_Snap is definded at form level


SQLQ = "SELECT WHSCC_USB.* , [WU_USB] - [WU_USBT] AS WU_USBO "
SQLQ = SQLQ & "FROM WHSCC_USB "
SQLQ = SQLQ & " ORDER BY WU_EFDATE,WU_ETDATE,WU_ORG"
  
Data1.RecordSource = SQLQ
Data1.Refresh

EEList = True
Exit Function

EEList_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "BankList", "WHSCC_USB", "Select")
Call RollBack '28July99 js

End Function

Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If


clpCode(0).Enabled = TF
dlpDateRange(0).Enabled = TF
dlpDateRange(1).Enabled = TF
medHours.Enabled = TF


End Sub

Function chkWHSCC()
Dim rsWHSCC As New ADODB.Recordset
Dim SQLQ As String, Msg As String, dd#
Dim X%, xID

On Error GoTo chkWHSCC_Err

chkWHSCC = False

If Len(clpCode(0)) < 1 Then
    MsgBox "Union Code is a required field"
    clpCode(0).SetFocus
    Exit Function
End If

If clpCode(0).Caption = "Unassigned" Then
    MsgBox "Union code must be valid"
    clpCode(0).SetFocus
    Exit Function
End If


For X% = 0 To 1
  If Len(dlpDateRange(X%)) > 0 Then
    If Not IsDate(dlpDateRange(X%)) Then
      MsgBox "Not a valid date"
      dlpDateRange(X%) = ""
      dlpDateRange(X%).SetFocus
      Exit Function
    End If
  Else
      MsgBox "Date is a required field"
      dlpDateRange(X%).SetFocus
      Exit Function
  End If
Next X%
  
If Not IsNumeric(medHours) Then
      MsgBox "Bank is not a number"
      medHours.SetFocus
      Exit Function
End If
If Not IsEmpty(Data1.Recordset("WU_ID")) Then
    xID = Data1.Recordset("WU_ID")
Else
    xID = -999
End If
SQLQ = "SELECT * FROM WHSCC_USB WHERE WU_EFDATE = ('" & Format(dlpDateRange(0), "mmm dd,yyyy") & "') "
SQLQ = SQLQ & "AND WU_ETDATE = ('" & Format(dlpDateRange(1), "mmm dd,yyyy") & "') "
SQLQ = SQLQ & "AND WU_ORG = '" & clpCode(0) & "' "
SQLQ = SQLQ & "AND WU_ID <> " & xID & " "
rsWHSCC.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsWHSCC.EOF Then
      MsgBox "Duplicate Union Code found within Sick Entitlement Date Range. "
      clpCode(0).SetFocus
      rsWHSCC.Close
      Exit Function
End If
rsWHSCC.Close

chkWHSCC = True

Exit Function

chkWHSCC_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkWHSCC", "WHSCC_USB", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub Display_Value()
    Dim SQLQ

    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
    SQLQ = "SELECT WHSCC_USB.* " ', [WU_USB] - [WU_USBT] AS WU_USBO "
    SQLQ = SQLQ & "FROM WHSCC_USB "
    SQLQ = SQLQ & "WHERE WU_ID = " & Data1.Recordset!WU_ID
    SQLQ = SQLQ & " ORDER BY WU_EFDATE,WU_ETDATE,WU_ORG"
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    valOuts = rsDATA("WU_USB") - IIf(IsNull(rsDATA("WU_USBT")), 0, rsDATA("WU_USBT"))
    Call cmdModify_Click
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT WHSCC_USB.* , [WU_USB] - [WU_USBT] AS WU_USBO "
        SQLQ = SQLQ & "FROM WHSCC_USB "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
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
RelateMode = NothingRelate
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_WHSCC_USB
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
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

End Sub
