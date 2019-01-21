VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEAddPayrollIDData 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Additional Payroll ID Data"
   ClientHeight    =   7965
   ClientLeft      =   105
   ClientTop       =   1035
   ClientWidth     =   9330
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
   ScaleHeight     =   7965
   ScaleWidth      =   9330
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtADPBranchNo 
      Appearance      =   0  'Flat
      DataField       =   "PY_ADP_BRANCH"
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
      Left            =   2235
      TabIndex        =   0
      Tag             =   "00-ADP Branch #"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.TextBox txtPayrollID 
      Appearance      =   0  'Flat
      DataField       =   "PY_PAYROLL_ID"
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
      Left            =   2235
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "00-Payroll ID"
      Top             =   3240
      Width           =   1815
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feADPPayroll.frx":0000
      Height          =   1845
      Left            =   180
      OleObjectBlob   =   "feADPPayroll.frx":0014
      TabIndex        =   4
      Top             =   660
      Width           =   8895
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "PY_DEPTNO"
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Tag             =   "01-Department"
      Top             =   3960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   1
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   7305
      Width           =   9330
      _Version        =   65536
      _ExtentX        =   16457
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
         Left            =   6540
         Top             =   105
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
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   7140
         Top             =   210
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PY_LDATE"
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
      Left            =   2670
      MaxLength       =   25
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PY_LTIME"
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
      Left            =   4470
      MaxLength       =   25
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PY_LUSER"
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
      Left            =   6150
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9330
      _Version        =   65536
      _ExtentX        =   16457
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
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
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
         Left            =   6360
         TabIndex        =   17
         Top             =   115
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         TabIndex        =   10
         Top             =   110
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         TabIndex        =   9
         Top             =   115
         Width           =   720
      End
   End
   Begin INFOHR_Controls.CodeLookup clpGLNbr 
      DataField       =   "PY_GLNO"
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "00-General Ledger - Code"
      Top             =   3600
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin VB.Label lblADPBranch 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ADP Branch #"
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
      Left            =   330
      TabIndex        =   19
      Top             =   2925
      Width           =   1095
   End
   Begin VB.Label lblPayrollID 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Payroll ID"
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
      Left            =   330
      TabIndex        =   18
      Top             =   3285
      Width           =   1095
   End
   Begin VB.Label lblGL 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ADP GL #"
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
      Left            =   330
      TabIndex        =   16
      Top             =   3630
      Width           =   735
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ADP Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   330
      TabIndex        =   14
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "PY_EMPNBR"
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
      Left            =   1710
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "PY_COMPNO"
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
      Left            =   30
      TabIndex        =   13
      Top             =   6000
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEAddPayrollIDData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AddChg
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fGLBNew
Dim RSDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Private Function chkEAddPayrollID()
Dim SQLQ As String, Msg As String, dd#
Dim flgADPDeptExists As Boolean

chkEAddPayrollID = False

On Error GoTo chkEComment_Err


If clpGLNbr.Caption = "Unassigned" Then
    MsgBox lStr("ADP GL # must be valid")
    clpGLNbr.SetFocus
    Exit Function
End If

If Len(clpDept.Text) < 1 Then
    MsgBox lStr("ADP Department is a required field")
    clpDept.SetFocus
    Exit Function
End If
 
If clpDept.Caption = "Unassigned" Then
    MsgBox lStr("ADP Department must be valid")
    clpDept.SetFocus
    Exit Function
End If

'Check if there is already same ADP Department in the table
flgADPDeptExists = ADP_Department_Exists
If flgADPDeptExists Then
    MsgBox lStr("Additional Payroll ID Data") & " for '" & clpDept.Text & "' " & lStr("ADP Department") & " already exists. Cannot save duplicate " & lStr("ADP Department") & " record.", vbExclamation, "Duplicate " & lStr("ADP Department")
    clpDept.SetFocus
    Exit Function
End If

chkEAddPayrollID = True

Exit Function

chkEComment_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEADPPayroll", "HR_PAYROLLID_DATA", "edit/Add")
Call RollBack '28July99 js

End Function

Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

fGLBNew = False
Call SET_UP_MODE

RSDATA.CancelUpdate
Call Display_Value

'Call ST_UPD_MODE(True)  ' reset screen's attributes

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_PAYROLLID_DATA", "Cancel")
Call RollBack '28July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()

Unload Me
If glbOnTop = "FRMEADDPAYROLLIDDATA" Then glbOnTop = ""

Call NextForm
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, X
 
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub


If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    RSDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fGLBNew = False

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_PAYROLLID_DATA", "Delete")
Call RollBack '28July99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
'clpCode(1).Enabled = True
'clpCode(1).SetFocus

AddChg = "C"

fGLBNew = False

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_PAYROLLID_DATA", "Modify")
Call RollBack '28July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String

fGLBNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

On Error GoTo AddN_Err

Call Set_Control("B", Me)

RSDATA.AddNew

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

lblCNum.Caption = "001"

AddChg = "A"
fGLBNew = True

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_PAYROLLID_DATA", "Add")
Call RollBack '28July99 js

End Sub

Sub cmdOK_Click()
Dim X, xID
Dim rsCOM As New ADODB.Recordset
On Error GoTo Add_Err

If Not chkEAddPayrollID() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, RSDATA)

If glbtermopen Then
    RSDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    RSDATA.Update
    gdbAdoIhr001X.CommitTrans
    RSDATA.Resync
    xID = RSDATA!PY_ID
Else
    gdbAdoIhr001.BeginTrans
    RSDATA.Update
    gdbAdoIhr001.CommitTrans
    RSDATA.Resync
    xID = RSDATA!PY_ID
End If
Data1.Refresh

Data1.Recordset.Find "PY_ID=" & xID
fGLBNew = False

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

Exit Sub

Add_Err:
If Err = 3022 Then
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PAYROLLID_DATA", "Update")
Call RollBack '28July99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String, dscGroup$

    RHeading = lblEEName & "'s " & lStr("Additional Payroll ID Data")
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
'
'    If Not glbtermopen Then
'        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgesaldist.rpt"
'        Me.vbxCrystal.SelectionFormula = "{HR_EMP_SALDIST.EB_EMPNBR} = " & glbLEE_ID
'        If glbSQL Or glbOracle Then
'            Me.vbxCrystal.Connect = RptODBC_SQL
'        Else
'            Me.vbxCrystal.Connect = "PWD=petman;"
'            Me.vbxCrystal.DataFiles(0) = glbIHRDB
'            Me.vbxCrystal.DataFiles(1) = glbIHRDB
'        End If
'    Else
'        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgesaldist.rpt"
'        Me.vbxCrystal.SelectionFormula = "{Term_EMP_SALDIST.TERM_SEQ} = " & glbTERM_Seq
'        If glbSQL Or glbOracle Then
'            Me.vbxCrystal.Connect = RptODBC_SQL
'        Else
'            Me.vbxCrystal.Connect = "PWD=petman;"
'            Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
'            Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
'        End If
'    End If
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
 '   cmdPrint.Enabled = True
    
End Sub

Sub cmdView_Click()
Dim RHeading As String, dscGroup$

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    RHeading = lblEEName & "'s " & lStr("Additional Payroll ID Data")
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    
'    If Not glbtermopen Then
'        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgesaldist.rpt"
'        Me.vbxCrystal.SelectionFormula = "{HR_EMP_SALDIST.EB_EMPNBR} = " & glbLEE_ID
'        If glbSQL Or glbOracle Then
'            Me.vbxCrystal.Connect = RptODBC_SQL
'        Else
'            Me.vbxCrystal.Connect = "PWD=petman;"
'            Me.vbxCrystal.DataFiles(0) = glbIHRDB
'            Me.vbxCrystal.DataFiles(1) = glbIHRDB
'        End If
'    Else
'        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgesaldist.rpt"
'        Me.vbxCrystal.SelectionFormula = "{Term_EMP_SALDIST.TERM_SEQ} = " & glbTERM_Seq
'        If glbSQL Or glbOracle Then
'            Me.vbxCrystal.Connect = RptODBC_SQL
'        Else
'            Me.vbxCrystal.Connect = "PWD=petman;"
'            Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
'            Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
'        End If
'    End If
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
   ' cmdPrint.Enabled = True
    
End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_PAYROLLID_DATA"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * from HR_PAYROLLID_DATA"
    SQLQ = SQLQ & " where PY_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY PY_ADP_BRANCH,PY_PAYROLL_ID,PY_GLNO,PY_DEPTNO"
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Additional Payroll Data Retrieve", "HR_PAYROLLID_DATA", "SELECT")
Call RollBack '28July99 js

Exit Function

End Function

Private Sub Form_Activate()
Call SET_UP_MODE
'Me.cmdModify_Click
    glbOnTop = "FRMEADDPAYROLLIDDATA"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEADDPAYROLLIDDATA"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEADDPAYROLLIDDATA"
AddChg = " "

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

Call setCaption(lblADPBranch)
Call setCaption(lblGL)
Call setCaption(lblDept)

If Len(glbLEE_SName) < 1 Then Exit Sub

Screen.MousePointer = HOURGLASS

Me.vbxTrueGrid.SetFocus
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = lStr("Additional Payroll ID Data") & " - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
 
lblEENum.Caption = ShowEmpnbr(lblEEID)

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
 '  cmdModify.Enabled = False
Else
 '  cmdModify.Enabled = True
   Data1.Recordset.MoveFirst
End If

Call Display_Value

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

If gSec_Upd_AddPayrollIDData Then
    Call ST_UPD_MODE(True)
'Else
'    Call ST_UPD_MODE(False)             '
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

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
    Set frmEAddPayrollIDData = Nothing 'carmen may 00
    Call NextForm
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

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF
clpDept.Enabled = TF
clpGLNbr.Enabled = TF
txtPayrollID.Enabled = TF
txtADPBranchNo.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'vbxTrueGrid.Enabled = FT

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
fUPMode = TF    ' update mode

End Sub

Private Sub txtADPBranchNo_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtPayrollID_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
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
    
    If glbtermopen Then         'Lucy July 5, 2000
        SQLQ = "Select * from Term_PAYROLLID_DATA"
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    Else
        SQLQ = "Select * from HR_PAYROLLID_DATA"
        SQLQ = SQLQ & " where PY_EMPNBR = " & glbLEE_ID
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
     
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err

'If Not Fnd_Match_Data1() Then Exit Sub 'MsgBox "No Records Found."
Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_PAYROLLID_DATA", "Add")
Call RollBack '28July99 js

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
Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        If glbtermopen Then
            RSDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Me.cmdModify_Click
       Exit Sub
    End If
    
    
If glbtermopen Then
    SQLQ = "Select * from Term_PAYROLLID_DATA"
    SQLQ = SQLQ & " WHERE PY_ID = " & Data1.Recordset!PY_ID
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HR_PAYROLLID_DATA"
    SQLQ = SQLQ & " where PY_ID = " & Data1.Recordset!PY_ID
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    
Call SET_UP_MODE
Me.cmdModify_Click

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fGLBNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fGLBNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_AddPayrollIDData
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
If fGLBNew Then
    UpdateState = NewRecord
    TF = True
ElseIf RSDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    Me.Caption = lStr("Additional Payroll ID Data") & " - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Function ADP_Department_Exists()
Dim rsPayrollID As New ADODB.Recordset
Dim SQLQ As String

    SQLQ = "SELECT * FROM HR_PAYROLLID_DATA"
    SQLQ = SQLQ & " WHERE PY_DEPTNO = '" & clpDept.Text & "'"
    If Not fGLBNew Then
        SQLQ = SQLQ & " AND PY_ID <> " & Data1.Recordset!PY_ID
    End If
    rsPayrollID.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsPayrollID.EOF Then
        ADP_Department_Exists = False
    Else
        ADP_Department_Exists = True
    End If
    rsPayrollID.Close
    Set rsPayrollID = Nothing
End Function
