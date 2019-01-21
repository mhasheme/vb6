VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmProductLineOperation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product Line / Operation"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "TB_DESC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2115
      MaxLength       =   50
      TabIndex        =   23
      Tag             =   "01-Code"
      Top             =   5520
      Width           =   4470
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   5910
      Width           =   9360
      _Version        =   65536
      _ExtentX        =   16510
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
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
         Left            =   5700
         TabIndex        =   8
         Tag             =   "Print Code Listing Report"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
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
         Left            =   4890
         TabIndex        =   7
         Tag             =   "Delete code listed above"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
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
         Left            =   4080
         TabIndex        =   6
         Tag             =   "Add a new Code"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
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
         Left            =   3225
         TabIndex        =   5
         Tag             =   "Cancel the changes made"
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
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
         Left            =   2415
         TabIndex        =   4
         Tag             =   "Save the changes made"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
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
         Left            =   1635
         TabIndex        =   3
         Tag             =   "Edit the Information"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
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
         Left            =   855
         TabIndex        =   2
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
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
         Left            =   75
         TabIndex        =   1
         Tag             =   "Select the Code listed above"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   5505
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         DiscardSavedData=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   480
         Left            =   4770
         Top             =   30
         Visible         =   0   'False
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   847
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
   Begin VB.TextBox txtFullCode 
      Appearance      =   0  'Flat
      DataField       =   "TB_FULLCODE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2115
      MaxLength       =   23
      TabIndex        =   21
      Tag             =   "01-Code"
      Top             =   5160
      Width           =   1710
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   180
      MaxLength       =   25
      TabIndex        =   13
      Text            =   "Ldate"
      Top             =   5970
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1740
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "LTime"
      Top             =   5970
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TB_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3300
      MaxLength       =   25
      TabIndex        =   11
      Text            =   "LUser"
      Top             =   5970
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtTable 
      Appearance      =   0  'Flat
      DataField       =   "TB_NAME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5970
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "TB_COMPNO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   5130
      MaxLength       =   3
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5970
      Visible         =   0   'False
      Width           =   255
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmProductLineOperation.frx":0000
      Height          =   3885
      Left            =   240
      OleObjectBlob   =   "frmProductLineOperation.frx":0014
      TabIndex        =   14
      Tag             =   "Codes Listings"
      Top             =   60
      Width           =   8595
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      DataField       =   "TB_DIV"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   15
      Top             =   4080
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpRegion 
      DataField       =   "TB_REGION"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   18
      Top             =   4440
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin INFOHR_Controls.CodeLookup clpSection 
      DataField       =   "TB_SECTION"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   4800
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMOP"
      MaxLength       =   12
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      Height          =   195
      Left            =   360
      TabIndex        =   24
      Top             =   5520
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Combined Code"
      Height          =   195
      Left            =   360
      TabIndex        =   22
      Top             =   5160
      Width           =   1125
   End
   Begin VB.Label txtSection 
      AutoSize        =   -1  'True
      Caption         =   "Home Operation #"
      Height          =   195
      Left            =   360
      TabIndex        =   19
      Top             =   4800
      Width           =   1305
   End
   Begin VB.Label txtRegion 
      AutoSize        =   -1  'True
      Caption         =   "Product Line"
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   4440
      Width           =   900
   End
   Begin VB.Label txtDiv 
      AutoSize        =   -1  'True
      Caption         =   "Facility"
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   4080
      Width           =   480
   End
End
Attribute VB_Name = "frmProductLineOperation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewRec%
Dim fglbUDMode As Integer
Dim fglbRSOld As String
Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc
Private Function chkMastTable()
Dim SQLQ As String, Msg$, Tabl As String, Ky As String
Dim snapTabs As New ADODB.Recordset

On Error GoTo chkMastTable_Err

chkMastTable = False

If Len(clpDiv) < 1 Then
    MsgBox lStr("Division is a required field")
    clpDiv.SetFocus
    Exit Function
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("If Division Entered - it must be known")
        clpDiv.SetFocus
        Exit Function
    End If
End If
If Len(clpRegion) < 1 Then
    MsgBox lStr("Region is a required field")
    clpRegion.SetFocus
    Exit Function
End If
If fglbNewRec% Then
    snapTabs.Open "SELECT * FROM LN_PROD WHERE TB_FULLCODE='" & txtFullCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If Not snapTabs.EOF Then
        MsgBox "Duplicate records"
        Exit Function
    End If
    snapTabs.Close
End If
chkMastTable = True

Exit Function

chkMastTable_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRTABLE", "HRTABL", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function





Private Sub cmbDWM_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub clpDIV_LostFocus()
txtFullCode = clpDiv & "-" & clpRegion
If clpSection <> "" Then txtFullCode = txtFullCode & "-" & clpSection
End Sub

Private Sub clpRegion_LostFocus()
txtFullCode = clpDiv & "-" & clpRegion
If clpSection <> "" Then txtFullCode = txtFullCode & "-" & clpSection
End Sub

Private Sub clpSection_LostFocus()
txtFullCode = clpDiv & "-" & clpRegion
If clpSection <> "" Then txtFullCode = txtFullCode & "-" & clpSection
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Can_Err

Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call ST_UPD_MODE(False)  ' reset screen's attributes

fglbNewRec% = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRTABL", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If




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
Dim rsEMP As New ADODB.Recordset
Dim SQLQ, Msg$

On Error GoTo DelErr

If Data1.Recordset.RecordCount < 2 Then
    MsgBox "You can not delete the last reference for this code"
    Exit Sub
End If
SQLQ = "SELECT ED_EMPNBR, ED_SURNAME, ED_DEPTNO FROM HREMP "
SQLQ = SQLQ & " WHERE ED_EMPNBR IN (SELECT SE_EMPNBR FROM LN_EMPSKL WHERE SE_FULLCODE='" & txtFullCode & "')"

rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic

If rsEMP.BOF And rsEMP.EOF Then
    GoTo Lok
Else
    Msg$ = lStr("Employee presently assigned to this Code")
    Msg$ = Msg$ & Chr(10) & ShowEmpnbr(rsEMP("ED_EMPNBR"))
    Msg$ = Msg$ & Chr(10) & rsEMP("ED_SURNAME")
    Msg$ = Msg$ & Chr(10) & "Delete aborted."
    MsgBox Msg$
    rsEMP.Close
    Exit Sub
End If

Lok:        'looks ok to me
rsEMP.Close
Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Exit Sub
DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRTable", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdDelete_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String
'If Len(txtFindKey) > 0 Then
'    If clpFindDIV = "" Then clpFindDIV = clpDIV
'    SQLQ = "TB_KEY >= '" & clpDIV & txtFindKey & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ
'    If Data1.Recordset.EOF Then
'        Data1.Refresh
'    Else
'        txtFindKey = ""
'    End If
'    clpFindDIV = ""
'    Exit Sub
'End If
'
'If Len(txtFindDesc) > 0 Then
'    SQLQ = "TB_DESC >= '" & txtFindDesc & "'"
'    Data1.Recordset.Requery
'    Data1.Recordset.Find SQLQ
'    If Data1.Recordset.EOF Then
'        Data1.Refresh
'    Else
'        txtFindDesc = ""
'    End If
'    clpFindDIV = ""
'    Exit Sub
'End If


End Sub

Private Sub cmdFind_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()


On Error GoTo Mod_Err
Call ST_UPD_MODE(True)
clpDiv.Enabled = False
clpRegion.Enabled = False
clpSection.Enabled = False
txtFullCode.Enabled = False
clpRegion.SetFocus

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdModify_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr


Call ST_UPD_MODE(True)
fglbNewRec% = True

glbCodeRef = True


Data1.Recordset.AddNew
If glbDIVCount = 1 Then clpDiv = glbSDIV
txtTable.Text = glbTabNam
txtComp.Text = glbCompNo

If clpDiv.Visible Then
    clpDiv.SetFocus
Else
    clpRegion.SetFocus
End If

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HRTABLE", "HRTABL", "add new")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub CmdNew_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim SQLQ As String
Dim strK

On Error GoTo OK_Err

glbCodeRef = True   'table entrie modified/added - forces refresh
                    ' at form level of codes/descriptions.

If Not chkMastTable() Then Exit Sub

Call UpdUStats(Me)
strK = clpDiv & "-" & clpRegion
If clpSection <> "" Then strK = strK & "-" & clpSection

Data1.Recordset("TB_FULLCODE") = strK
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Data1.Recordset.Requery
Data1.Recordset.Find "TB_FULLCODE = '" & strK & "'"
Call ST_UPD_MODE(False)
fglbNewRec% = False


Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRTABL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.ReportTitle = "Table Codes for - " & glbTabNam
Me.vbxCrystal.BoundReportHeading = frmMTABLin.Caption
Me.vbxCrystal.WindowTitle = frmMTABLin.Caption & " Report"
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()


If Data1.Recordset.EOF And Data1.Recordset.BOF Then
  Exit Sub
End If

If Len(Data1.Recordset("TB_FULLCODE")) > 0 Then
    glbCode = Data1.Recordset("TB_FULLCODE")
    'glbCodeDesc = clpRegion & "," & clpSection
    If IsNull(Data1.Recordset("TB_DESC")) Then
        glbCodeDesc = ""
    Else
        glbCodeDesc = Data1.Recordset("TB_DESC")
    End If
    Unload Me
Else
    Exit Sub
End If


End Sub

Private Sub cmdSelect_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "data1 error", "HRTABL", "SELECT")


End Sub

Private Sub Form_Load()
Dim SQLQ As String
Dim VS
Screen.MousePointer = HOURGLASS
glbOnTop = "FRMPRODUCTLINEOPERATION"
glbCodeRef = False  'table entrie modified/added false
     

SQLQ = "SELECT * FROM LN_PROD WHERE TB_DIV in " & glbDIVList
SQLQ = SQLQ & " ORDER BY TB_DIV,TB_REGION,TB_SECTION"

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ
Data1.Refresh
VS = True
vbxTrueGrid.Columns(0).Visible = VS
glbCode = ""    'set to null - implies none found/cancel
glbCodeDesc = ""
Call INI_Controls(Me)

Call ST_UPD_MODE(False)

Screen.MousePointer = DEFAULT

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

fglbUDMode = TF     'in update/new mode

cmdOK.Enabled = TF
cmdCancel.Enabled = TF

cmdModify.Enabled = FT
cmdClose.Enabled = FT
cmdNew.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT
'cmdFind.Enabled = FT
cmdSelect.Enabled = FT


clpDiv.Enabled = TF
clpRegion.Enabled = TF
clpSection.Enabled = TF

txtFullCode.Enabled = False
txtDesc.Enabled = TF


'clpFindDIV.Enabled = FT
'txtFindKey.Enabled = FT
'txtFindDesc.Enabled = FT
vbxTrueGrid.Enabled = FT
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    cmdSelect.Enabled = False
    cmdModify.Enabled = False
    'cmdFind.Enabled = False
    cmdDelete.Enabled = False
End If
On Error GoTo ERR_EXIT
If Not gSec_Upd_Productline_Operation Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If
ERR_EXIT:
If Err.Number = 5 Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If
End Sub

Private Sub txtDesc_GotFocus()
Call SetPanHelp(ActiveControl)
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



Private Sub txtKey_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtShowKey_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtShowKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub



Private Sub vbxTrueGrid_DblClick()

If Not Me.vbxTrueGrid.EditActive Then
    If Not (Data1.Recordset.BOF Or Data1.Recordset.EOF) Then
        glbCode = Data1.Recordset("TB_FULLCODE")
        'glbCodeDesc = clpRegion & "," & clpSection
        If IsNull(Data1.Recordset("TB_DESC")) Then
            glbCodeDesc = ""
        Else
            glbCodeDesc = Data1.Recordset("TB_DESC")
        End If
    End If
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
        
        SQLQ = "SELECT * FROM LN_PROD WHERE TB_DIV in " & glbDIVList
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub
