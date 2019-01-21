VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUYTDBTI 
   Caption         =   "Year End Reduction For BD"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   9192
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   9192
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdYearEnd 
      Caption         =   "Year End Calculate"
      Height          =   375
      Left            =   3720
      TabIndex        =   25
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton cmdAverage 
      Caption         =   "Average Calculate"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   6720
      Width           =   2055
   End
   Begin VB.TextBox txtApplied 
      Appearance      =   0  'Flat
      DataField       =   "BT_APPLIED"
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   13
      Tag             =   "00-Applied"
      Top             =   5760
      Width           =   420
   End
   Begin VB.TextBox txtApprove 
      Appearance      =   0  'Flat
      DataField       =   "BT_APPROV"
      DataSource      =   " "
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      MaxLength       =   50
      TabIndex        =   12
      Tag             =   "00-Approved By"
      Top             =   5400
      Width           =   4020
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BT_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7470
      MaxLength       =   8
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BT_LDATE"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   6720
      MaxLength       =   12
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3345
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BT_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8160
      MaxLength       =   10
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   645
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fuytdbti.frx":0000
      Height          =   2535
      Left            =   120
      OleObjectBlob   =   "fuytdbti.frx":0014
      TabIndex        =   1
      Top             =   120
      Width           =   8775
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7080
      Top             =   5880
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3408
      _ExtentY        =   572
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
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7560
      Top             =   5400
      _ExtentX        =   593
      _ExtentY        =   593
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
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "BT_DIV"
      Height          =   285
      Left            =   2085
      TabIndex        =   5
      Tag             =   "00-Specific Division Desired"
      Top             =   2880
      Width           =   3195
      _ExtentX        =   5630
      _ExtentY        =   508
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
      Object.Height          =   288
      Enabled         =   0   'False
   End
   Begin MSMask.MaskEdBox medYear 
      DataField       =   "BT_YEAR"
      Height          =   285
      Left            =   2400
      TabIndex        =   6
      Tag             =   "00-Specific Year Desired"
      Top             =   3240
      Width           =   855
      _ExtentX        =   1503
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTotEmp 
      DataField       =   "BT_EMPNUM"
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Tag             =   "00-Specific Total Employees Desired"
      Top             =   3600
      Width           =   1530
      _ExtentX        =   2709
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medSysABS 
      DataField       =   "BT_SYSABS_AVG"
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Tag             =   "00-System Absent Average"
      Top             =   3960
      Width           =   1530
      _ExtentX        =   2709
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medUpdABS 
      DataField       =   "BT_UPDABS_AVG"
      Height          =   285
      Left            =   2400
      TabIndex        =   9
      Tag             =   "00-Updated Absent Average"
      Top             =   4320
      Width           =   1530
      _ExtentX        =   2709
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medSysLLE 
      DataField       =   "BT_SYSLLE_AVG"
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Tag             =   "00-System L/LE Average"
      Top             =   4680
      Width           =   1530
      _ExtentX        =   2709
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medUpdLLE 
      DataField       =   "BT_UPDLLE_AVG"
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Tag             =   "00-Updated L/LE Average"
      Top             =   5040
      Width           =   1530
      _ExtentX        =   2709
      _ExtentY        =   508
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,##0.00;(#,##0.00)"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Approved By"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   1425
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Applied"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   22
      Top             =   5760
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Employees"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "System Absent Average"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   20
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   2880
      Width           =   840
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Updated Absent Average"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   4320
      Width           =   1905
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "System L/LE Average"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   16
      Top             =   4680
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Updated L/LE Average"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   1785
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "BT_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7320
      TabIndex        =   14
      Top             =   3960
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lbl1 
      Caption         =   "BTI Only"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   7320
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmUYTDBTI"
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
Dim rsDATA As New ADODB.Recordset
Dim UpdateState As UpdateStateEnum
Dim SQLQ

Sub cmdClose_Click()
    Unload Me
End Sub
Sub cmdCancel_Click()

On Error GoTo Can_Err
fglbNew = False
If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If
fglbEmptyNew = False

rsDATA.CancelUpdate
Call Display_Value

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_BTI_AVERAGE", "Cancel")
Call RollBack '09June99 js

End Sub
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_BTI_AVERAGE", "Delete")
Call RollBack '09June99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)
medYear.SetFocus
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
clpDiv.Text = "BD"
txtApprove.Text = glbUserNAME
txtApplied.Text = "N"
lblCNum.Caption = "001"

fglbNew = True
Me.vbxTrueGrid.Enabled = False
fglbEmptyNew = True
Call SET_UP_MODE

medYear.SetFocus
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_BTI_AVERAGE", "Add")
Call RollBack '09June99 js

End Sub

Sub cmdOK_Click()
Dim x%

On Error GoTo cmdOK_Err

If Not chkUYTDBTI() Then Exit Sub


Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh

fglbNew = False
fglbEmptyNew = False
Call SET_UP_MODE

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_BTI_AVERAGE", "Update")
Call RollBack '09June99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String
'
'RHeading = "Payroll Matrix"
'Me.vbxCrystal.WindowTitle = RHeading & " Report"
'Me.vbxCrystal.BoundReportHeading = RHeading
''Me.vbxCrystal.Password = gstrAccPWord$
''Me.vbxCrystal.UserName = gstrAccUID$
'Me.vbxCrystal.Action = 1

End Sub
Sub cmdView_Click()
Dim RHeading As String
'
'RHeading = "Payroll Matrix"
'Me.vbxCrystal.WindowTitle = RHeading & " Report"
'Me.vbxCrystal.BoundReportHeading = RHeading
''Me.vbxCrystal.Password = gstrAccPWord$
''Me.vbxCrystal.UserName = gstrAccUID$
'Me.vbxCrystal.Destination = 0
'Me.vbxCrystal.Action = 1

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
UpdateRight = True 'gSec_Matrix
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

Private Sub clpDiv_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdAverage_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Calculate"
    Exit Sub
End If
Msg = ""
If Len(medTotEmp) > 0 Then
    Msg = Msg & "This Calculation has been done before. " & Chr(10)
End If
Msg = Msg & "Are You Sure You Want To Do "
Msg = Msg & "The Average Calculate For This Record?"
a% = MsgBox(Msg, 36, "Confirm ")

If a% <> 6 Then Exit Sub

Call Calcu_Average

Data1.Refresh

End Sub
Private Sub Calcu_Reduction_BD()
Dim rsTAtt As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsMain As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ, xNum, xCode, xSec, xEmpNo
Dim xUnexFlag, xEmlFlag, xUnexVal, xExcuVal, xEmlVal, glbDiv, glbYear, xYear, glbPointType
Dim I, xNTot, xDOA
Dim xFDate, xTDate, xABS_Upd, xLLE_Upd, xTVal
    xABS_Upd = Val(medUpdABS)
    xLLE_Upd = Val(medUpdLLE)
    xFDate = CVDate(GetMonth("Jan") & " 1," & medYear + 1)
    xTDate = CVDate(GetMonth("Dec") & " 31," & medYear + 1)
    xYear = Val(medYear)
    SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY' AND (ED_DIV = 'BD')"
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsMain.EOF Then
        I = 0
        xNTot = rsMain.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    Do While Not rsMain.EOF
        DoEvents
        MDIMain.panHelp(0).FloodPercent = (I / xNTot) * 100: I = I + 1
        xEmpNo = rsMain("ED_EMPNBR")
        'Get glbDiv
        glbDiv = rsMain("ED_DIV")
        xSec = rsMain("ED_SECTION")

        'Get the L/LE Point for this employee - Begin
        If xLLE_Upd > 0 Then
            'Get Carryover points first, if no carryover and then no reduction, because the balance can't be negative
            SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_LEPOINT FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
            SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xFDate) & " "
            SQLQ = SQLQ & "AND (AD_REASON='2100') "
            SQLQ = SQLQ & "AND AD_LEPOINT <>0 "
            rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTAtt.EOF Then
                If Not IsNull(rsTAtt("AD_LEPOINT")) Then
                    If rsTAtt("AD_LEPOINT") > 0 Then '= xLLE_Upd Then
                        xTVal = xLLE_Upd
                        If rsTAtt("AD_LEPOINT") < xLLE_Upd Then
                            xTVal = rsTAtt("AD_LEPOINT")
                        End If
                        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                        SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xFDate) & " "
                        SQLQ = SQLQ & "AND (AD_REASON='2400') "
                        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsTemp.EOF Then
                            rsTemp.AddNew
                            rsTemp("AD_COMPNO") = "001"
                            rsTemp("AD_EMPNBR") = xEmpNo
                            rsTemp("AD_DOA") = xFDate
                            rsTemp("AD_REASON") = "2400"
                            rsTemp("AD_HRS") = 0
                            rsTemp("AD_LEPOINT") = -xTVal 'xLLE_Upd
                            rsTemp("AD_INDICATOR") = 1
                            rsTemp("AD_SEN") = 0
                            rsTemp("AD_LDATE") = Date
                            rsTemp("AD_LTIME") = Time$
                            rsTemp("AD_LUSER") = glbUserID
                            rsTemp.Update
                            'Debug.Print xEmpNo
                        Else
                            rsTemp("AD_LEPOINT") = -xTVal 'xLLE_Upd
                            rsTemp("AD_INDICATOR") = 1
                            rsTemp("AD_SEN") = 0
                            rsTemp("AD_LDATE") = Date
                            rsTemp("AD_LTIME") = Time$
                            rsTemp("AD_LUSER") = glbUserID
                            rsTemp.Update
                        End If
                        rsTemp.Close
                    End If
                End If
            End If
            rsTAtt.Close
        End If
        'Get the L/LE Point for this employee - End
        
        'Get the ABS Point for this employee - Begin
        If xABS_Upd > 0 Then
            'Get Carryover points first, if no carryover and then no reduction, because the balance can't be negative
            SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_POINT FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
            SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xFDate) & " "
            SQLQ = SQLQ & "AND (AD_REASON='2200') "
            SQLQ = SQLQ & "AND AD_POINT <>0 "
            rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTAtt.EOF Then
                If Not IsNull(rsTAtt("AD_POINT")) Then
                    If rsTAtt("AD_POINT") > 0 Then '= xABS_Upd Then
                        xTVal = xABS_Upd
                        If rsTAtt("AD_POINT") < xABS_Upd Then
                            xTVal = rsTAtt("AD_POINT")
                        End If
                        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                        SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xFDate) & " "
                        SQLQ = SQLQ & "AND (AD_REASON='2300') "
                        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsTemp.EOF Then
                            rsTemp.AddNew
                            rsTemp("AD_COMPNO") = "001"
                            rsTemp("AD_EMPNBR") = xEmpNo
                            rsTemp("AD_DOA") = xFDate
                            rsTemp("AD_REASON") = "2300"
                            rsTemp("AD_HRS") = 0
                            rsTemp("AD_POINT") = -xTVal 'xABS_Upd
                            rsTemp("AD_INDICATOR") = 1
                            rsTemp("AD_SEN") = 0
                            rsTemp("AD_LDATE") = Date
                            rsTemp("AD_LTIME") = Time$
                            rsTemp("AD_LUSER") = glbUserID
                            rsTemp.Update
                            'Debug.Print xEmpNo
                        Else
                            rsTemp("AD_POINT") = -xTVal 'xABS_Upd
                            rsTemp("AD_INDICATOR") = 1
                            rsTemp("AD_SEN") = 0
                            rsTemp("AD_LDATE") = Date
                            rsTemp("AD_LTIME") = Time$
                            rsTemp("AD_LUSER") = glbUserID
                            rsTemp.Update
                        End If
                        rsTemp.Close
                    End If
                End If
            End If
            rsTAtt.Close
        End If
        'Get the ABS Point for this employee - End
        
        rsMain.MoveNext
    Loop
    
    'Modify the Yearly Average Table
    SQLQ = "SELECT * FROM HR_BTI_AVERAGE WHERE BT_DIV= '" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND BT_YEAR = '" & medYear.Text & "' "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        rsEmp("BT_APPLIED") = "Y"
        rsEmp.Update
    End If
    rsEmp.Close
    
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = DEFAULT
    MsgBox "Update completed"
End Sub
Private Sub Calcu_Average()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ, fDATE, tDATE, xYear
Dim TotEmp, TotAbs, TotLLE, AveAbs, AveLLE
    TotEmp = 0: TotAbs = 0: TotLLE = 0:
    fDATE = CVDate(GetMonth("Jan") & " 1," & medYear)
    tDATE = CVDate(GetMonth("Dec") & " 31," & medYear)
    xYear = medYear
    'For Active Emp ================================================================
    'Count Total Employee
    SQLQ = "SELECT COUNT('ED_EMPNBR') AS TOTEMP FROM HREMP WHERE ED_SECTION='HRLY' "
    SQLQ = SQLQ & "AND ED_DIV='" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND ED_DOH <=" & Date_SQL(tDATE) & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If rsEmp("TOTEMP") > 0 Then
            TotEmp = rsEmp("TOTEMP")
        End If
    End If
    rsEmp.Close
    
    'Get the ABS Point  - Begin
    SQLQ = "SELECT SUM(AD_POINT) AS TOTNUM FROM HR_ATTENDANCE, HREMP "
    SQLQ = SQLQ & "WHERE HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR  "
    SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & xYear & " "
    SQLQ = SQLQ & "AND AD_POINT <>0 "
    SQLQ = SQLQ & "AND ED_SECTION='HRLY' "
    SQLQ = SQLQ & "AND ED_DIV='" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND ED_DOH <=" & Date_SQL(tDATE) & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If rsEmp("TOTNUM") > 0 Then
            TotAbs = rsEmp("TOTNUM")
        End If
    End If
    rsEmp.Close
    
    'Get the LLE Point  - Begin
    SQLQ = "SELECT SUM(AD_LEPOINT) AS TOTNUM FROM HR_ATTENDANCE, HREMP "
    SQLQ = SQLQ & "WHERE (HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR)  "
    SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & xYear & " "
    SQLQ = SQLQ & "AND AD_LEPOINT <>0 "
    SQLQ = SQLQ & "AND ED_SECTION='HRLY' "
    SQLQ = SQLQ & "AND ED_DIV='" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND ED_DOH <=" & Date_SQL(tDATE) & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If rsEmp("TOTNUM") > 0 Then
            TotLLE = rsEmp("TOTNUM")
        End If
    End If
    rsEmp.Close
    
    'For Term Emp ================================================================
    'Count Total Employee '
    SQLQ = "SELECT COUNT('ED_EMPNBR') AS TOTEMP FROM Term_HREMP INNER JOIN Term_HRTRMEMP "
    SQLQ = SQLQ & "ON Term_HRTRMEMP.TERM_SEQ = Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & "WHERE ED_SECTION='HRLY' "
    SQLQ = SQLQ & "AND ED_DIV='" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND ED_DOH <=" & Date_SQL(tDATE) & " "
    SQLQ = SQLQ & "AND (Term_HRTRMEMP.Term_DOT > " & Date_SQL(tDATE) & ") "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If rsEmp("TOTEMP") > 0 Then
            TotEmp = TotEmp + rsEmp("TOTEMP")
        End If
    End If
    rsEmp.Close
    'Get the ABS Point  - Begin
    SQLQ = "SELECT SUM(AD_POINT) AS TOTNUM FROM Term_ATTENDANCE,Term_HREMP,Term_HRTRMEMP "
    SQLQ = SQLQ & "WHERE Term_ATTENDANCE.TERM_SEQ=Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & "AND Term_HREMP.TERM_SEQ=Term_HRTRMEMP.TERM_SEQ "
    SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & xYear & " "
    SQLQ = SQLQ & "AND AD_POINT <>0 "
    SQLQ = SQLQ & "AND ED_SECTION='HRLY' "
    SQLQ = SQLQ & "AND ED_DIV='" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND ED_DOH <=" & Date_SQL(tDATE) & " "
    SQLQ = SQLQ & "AND Term_DOT >" & Date_SQL(tDATE) & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If rsEmp("TOTNUM") > 0 Then
            TotAbs = TotAbs + rsEmp("TOTNUM")
        End If
    End If
    rsEmp.Close
    'Get the LLE Point  - Begin
    SQLQ = "SELECT SUM(AD_LEPOINT) AS TOTNUM FROM Term_ATTENDANCE,Term_HREMP,Term_HRTRMEMP "
    SQLQ = SQLQ & "WHERE Term_ATTENDANCE.TERM_SEQ=Term_HREMP.TERM_SEQ "
    SQLQ = SQLQ & "AND Term_HREMP.TERM_SEQ=Term_HRTRMEMP.TERM_SEQ "
    SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & xYear & " "
    SQLQ = SQLQ & "AND AD_LEPOINT <>0 "
    SQLQ = SQLQ & "AND ED_SECTION='HRLY' "
    SQLQ = SQLQ & "AND ED_DIV='" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND ED_DOH <=" & Date_SQL(tDATE) & " "
    SQLQ = SQLQ & "AND Term_DOT >" & Date_SQL(tDATE) & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If rsEmp("TOTNUM") > 0 Then
            TotLLE = TotLLE + rsEmp("TOTNUM")
        End If
    End If
    rsEmp.Close
    
    
    'Modify the Yearly Average Table
    SQLQ = "SELECT * FROM HR_BTI_AVERAGE WHERE BT_DIV= '" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND BT_YEAR = '" & medYear.Text & "' "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF And TotEmp > 0 Then
        rsEmp("BT_EMPNUM") = TotEmp 'AveAbs, AveLLE
        'Abs
        AveAbs = Round((TotAbs / TotEmp), 2)
        rsEmp("BT_SYSABS_AVG") = AveAbs
        AveAbs = Round((TotAbs / TotEmp), 0)
        If AveAbs > 4 Then AveAbs = 4
        rsEmp("BT_UPDABS_AVG") = AveAbs
        'LLE
        AveLLE = Round((TotLLE / TotEmp), 2)
        rsEmp("BT_SYSLLE_AVG") = AveLLE
        AveLLE = Round((TotLLE / TotEmp), 0)
        If AveLLE > 2 Then AveLLE = 2
        rsEmp("BT_UPDLLE_AVG") = AveLLE
        rsEmp.Update
    End If
    rsEmp.Close
    Screen.MousePointer = DEFAULT
    MsgBox "Update completed"
    
End Sub

Private Sub cmdYearEnd_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Calculate"
    Exit Sub
End If
If Not IsNumeric(medUpdABS) Then
    MsgBox "Please do Average Calculate first."
    Exit Sub
End If
Msg = ""
If (txtApplied) = "Y" Then
    Msg = Msg & "This Calculation has been done before. " & Chr(10)
End If
Msg = Msg & "Are You Sure You Want To Do "
Msg = Msg & "The Average Calculate For This Record?"
a% = MsgBox(Msg, 36, "Confirm ")

If a% <> 6 Then Exit Sub

Call Calcu_Reduction_BD

Data1.Refresh

End Sub

Private Sub Form_Activate()
'Me.cmdModify_Click
Call SET_UP_MODE
End Sub

Private Sub Form_Load()

Me.Caption = glbFormCaption
Me.Show
glbOnTop = "frmUYTDBTI"

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HR_BTI_AVERAGE ORDER BY BT_DIV,BT_YEAR"
Data1.Refresh

Call setRptCaption(Me)
Screen.MousePointer = DEFAULT
Call ST_UPD_MODE(False)
                                                
vbxTrueGrid.Columns(0).Caption = lStr("Division")

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
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
'clpDiv.Enabled = TF
medYear.Enabled = TF
medTotEmp.Enabled = TF
medSysABS.Enabled = TF
medUpdABS.Enabled = TF
medSysLLE.Enabled = TF
medUpdLLE.Enabled = TF
'txtApprove.Enabled = TF
'txtApplied.Enabled = TF

End Sub

Private Sub medSysABS_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medSysLLE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTarget_Change()

End Sub

Private Sub medTarget_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTotEmp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medUpdABS_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medUpdLLE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtApplied_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtApprove_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HR_BTI_AVERAGE "
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "Year End Reduction", "Select")
Call RollBack
End Sub

Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_BTI_AVERAGE WHERE BT_ID= " & Data1.Recordset!BT_ID
    
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    Call SET_UP_MODE
End Sub

Private Function chkUYTDBTI()
Dim rsTemp As New ADODB.Recordset
Dim Msg As String
Dim x%, xchk, SQLQ

chkUYTDBTI = False

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

If Not IsNumeric(medYear) Then
    MsgBox ("Invalid Year")
    medYear.SetFocus
    Exit Function
End If
If Len(medYear) <> 4 Then
    MsgBox ("Not 4 Digit Year")
    medYear.SetFocus
    Exit Function
End If
If fglbNew Then
    SQLQ = "SELECT * FROM HR_BTI_AVERAGE WHERE BT_DIV= '" & clpDiv.Text & "' "
    SQLQ = SQLQ & "AND BT_YEAR = '" & medYear.Text & "' "
    'SQLQ = SQLQ & "AND BT_ID <> " & Data1.Recordset!BT_ID & " "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If rsTemp.BOF And rsTemp.EOF Then
        rsTemp.Close
    Else
        Msg = ("This record is duplicate")
        MsgBox Msg
        medYear.SetFocus
        Exit Function
    End If
End If

xchk = False


chkUYTDBTI = True

End Function
