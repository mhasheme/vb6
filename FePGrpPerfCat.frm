VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPGrpPerfCatLnk 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Position Group and Performance Category Link"
   ClientHeight    =   5565
   ClientLeft      =   1485
   ClientTop       =   885
   ClientWidth     =   7395
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
   ScaleHeight     =   5565
   ScaleWidth      =   7395
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   7080
      Top             =   3840
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   7
      Top             =   4905
      Width           =   7395
      _Version        =   65536
      _ExtentX        =   13044
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
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   6720
         TabIndex        =   8
         Tag             =   "Select Province listed above"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Tag             =   "Close and exit screen"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1035
         TabIndex        =   10
         Tag             =   "Edit the information above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1845
         TabIndex        =   11
         Tag             =   "Save the changes made"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3525
         TabIndex        =   13
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4335
         TabIndex        =   14
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5745
         TabIndex        =   15
         Tag             =   "Print the Province listing report"
         Top             =   165
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   5235
      TabIndex        =   5
      Tag             =   "Find specific record"
      Top             =   4890
      Width           =   735
   End
   Begin VB.TextBox txtFindDesc 
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
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Tag             =   "00-Search Description"
      Top             =   4920
      Width           =   3975
   End
   Begin VB.TextBox txtFindKey 
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
      Height          =   285
      Left            =   360
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "00-Search Code"
      Top             =   4920
      Width           =   540
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "PP_COMPNO"
      DataSource      =   "Data1"
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
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "001"
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "FePGrpPerfCat.frx":0000
      Height          =   3105
      Left            =   120
      OleObjectBlob   =   "FePGrpPerfCat.frx":0014
      TabIndex        =   0
      Tag             =   "Province Listings"
      Top             =   120
      Width           =   7095
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7080
      Top             =   4440
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
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PJ_GRPCD"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2460
      TabIndex        =   1
      Tag             =   "01-Position Group Code"
      Top             =   3480
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PJ_CATECODE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2460
      TabIndex        =   2
      Tag             =   "00-Performance Category - Code"
      Top             =   3960
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPG"
   End
   Begin Threed.SSCheck chkInclFriesensForms 
      DataField       =   "PJ_FRIESENS_FORMS"
      DataSource      =   "Data1"
      Height          =   225
      Left            =   120
      TabIndex        =   18
      Tag             =   "Include this code Friesens Forms?"
      Top             =   4440
      Width           =   2850
      _Version        =   65536
      _ExtentX        =   5027
      _ExtentY        =   397
      _StockProps     =   78
      Caption         =   "Include in Friesens Forms                  "
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3525
      Width           =   2355
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Performance Category"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   4005
      Width           =   2355
   End
End
Attribute VB_Name = "frmPGrpPerfCatLnk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew As Integer
Dim fglbNewRec%, xOldCode As String
Dim fglbID As Integer
Dim xLinkItem As String
Dim xPosGrp, xPerfCat As String

Private Function chkPosGrp_PerfCat()
Dim Prov$, Msg$
Dim rsPosGrpCat As New ADODB.Recordset
Dim SQLQ As String

chkPosGrp_PerfCat = False

On Error GoTo chkPosGrp_PerfCat_Err

If Len(clpCode(0)) < 1 Then
    MsgBox lblTitle(0).Caption & " is a required field"
    clpCode(0).SetFocus
    Exit Function
End If
If Len(clpCode(1)) < 1 Then
    MsgBox lblTitle(1).Caption & " is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

'Check if this combination of Position Group and Performance Category already exists
Set rsPosGrpCat = Nothing
SQLQ = "SELECT * FROM HR_PERF_JOBGRP"
SQLQ = SQLQ & " WHERE PJ_GRPCD = '" & clpCode(0).Text & "'"
SQLQ = SQLQ & " AND PJ_CATECODE = '" & clpCode(1).Text & "'"
rsPosGrpCat.Open SQLQ, gdbAdoIhr001, adOpenStatic
If fglbNewRec% = True Then
    If Not rsPosGrpCat.EOF Then
        'Combination already exist
        MsgBox "This " & lblTitle(0).Caption & " and " & lblTitle(1).Caption & " link already exists."
        clpCode(0).SetFocus
        Exit Function
    End If
    rsPosGrpCat.Close
Else
    If (xPosGrp <> clpCode(0).Text) Or (xPerfCat <> clpCode(1).Text) Then
        If Not rsPosGrpCat.EOF Then
            'Combination already exist
            MsgBox "This " & lblTitle(0).Caption & " and " & lblTitle(1).Caption & " link already exists."
            clpCode(0).SetFocus
            Exit Function
        End If
        rsPosGrpCat.Close
    End If
End If

chkPosGrp_PerfCat = True

Exit Function

chkPosGrp_PerfCat_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HR_PERF_JOBGRP", "chkPosGrp_PerfCat")
Resume Next

End Function

Private Sub chkInclFriesensForms_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdCancel_Click()
Dim bk
'On Error GoTo Can_Err

Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Data1.Refresh
End If

Call modSTUPD(False)  ' reset screen's attributes

cmdClose.SetFocus

fglbNewRec% = False
xPosGrp = ""
xPerfCat = ""

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_PERF_JOBGRP", "Cancel")
Resume Next

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Msg As String, a%

On Error GoTo DelErr

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Exit Sub
End If

Data1.Recordset.Delete

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

Data1.Refresh

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If

Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Single", "Delete")
Call RollBack '09June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdFind_Click()
'Dim SQLQ$
'
'txtFindKey.SetFocus 'added by Marlon Cowan 9/16/97
'
'If Len(txtFindKey) > 0 Then
'    SQLQ$ = "CODE >= '" & txtFindKey.Text & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ$
'    If Data1.Recordset.EOF Then
'        Data1.Refresh
'    Else
'        txtFindKey = ""
'    End If
'    Exit Sub
'End If
'
'If Len(txtFindDesc) > 0 Then
'    SQLQ$ = "DESCR >= '" & txtFindDesc.Text & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ$
'    If Data1.Recordset.EOF Then
'        Data1.Refresh
'    Else
'        txtFindDesc = ""
'    End If
'    Exit Sub
'End If
'
'End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
xPosGrp = clpCode(0).Text
xPerfCat = clpCode(1).Text

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '09June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr

glbCodeRef = True

Data1.Recordset.AddNew

txtComp.Text = glbCompNo
xPosGrp = ""
xPerfCat = ""

fglbNewRec% = True

Call modSTUPD(True)

clpCode(0).SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_PERF_JOBGRP", "AddNew")
Resume Next

End Sub

Private Sub cmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim Desc As String
Dim ProvCode

On Error GoTo OK_Err

If Not chkPosGrp_PerfCat() Then Exit Sub

Data1.Recordset("PJ_COMPNO") = txtComp
Data1.Recordset("PJ_LDATE") = Format(Now, "SHORT DATE")
Data1.Recordset("PJ_LTIME") = Time$
Data1.Recordset("PJ_LUSER") = glbUserID
'Data1.Recordset("PJ_FRIESENS_FORMS") = chkInclFriesensForms
Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    fglbID = Data1.Recordset("PJ_ID")
    Data1.Refresh
    Data1.Recordset.Find "PJ_ID=" & fglbID & " "
End If

fglbNewRec% = False

Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption

Resume Next
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PERF_JOBGRP", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Province/State Codes"
Me.vbxCrystal.WindowTitle = "Position Group and Performance Category Link Report"
Me.vbxCrystal.BoundReportHeading = "Position Group/Peformance Category"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'
'glbFrmCaption$ = Me.Caption
'glbErrNum& = ErrorNumber
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_OHS_RLINK_EVENT", "SELECT")
'
'End Sub

Private Sub Form_Load()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT * FROM HR_PERF_JOBGRP ORDER BY PJ_GRPCD,PJ_CATECODE "
    Data1.Refresh

    Call modSTUPD(False)            'Jaddy 10/18/99
    
    Call INI_Controls(Me)
    
    Screen.MousePointer = DEFAULT   '
                                
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

cmdOK.Enabled = TF          'May99 js
cmdCancel.Enabled = TF      '
cmdClose.Enabled = FT
cmdPrint.Enabled = FT       '
cmdFind.Enabled = FT        '
cmdSelect.Enabled = FT
cmdDelete.Enabled = FT
If gSec_Upd_Performance And gSec_Upd_Job_Master Then  '
    cmdModify.Enabled = FT      '
    cmdNew.Enabled = FT         '
    cmdDelete.Enabled = FT      '
Else
    cmdModify.Enabled = False      '
    cmdNew.Enabled = False   '
    cmdDelete.Enabled = False
End If

If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
chkInclFriesensForms.Enabled = TF

txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
vbxTrueGrid.Enabled = FT

'If glbDivInhSel Then
'    cmdSelect.Enabled = False
'End If

End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
                       
        SQLQ = "SELECT * FROM HR_PERF_JOBGRP "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag

        Data1.RecordSource = SQLQ
        Data1.Refresh

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

'Public Property Let LinkItem(vData As String)
'    xLinkItem = vData
'End Property
'
'Public Property Get LinkItem() As String
'    LinkItem = xLinkItem
'End Property

