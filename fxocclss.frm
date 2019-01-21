VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOCCLASS 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NOC Codes"
   ClientHeight    =   6120
   ClientLeft      =   1260
   ClientTop       =   690
   ClientWidth     =   7140
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6120
   ScaleWidth      =   7140
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4200
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      TabIndex        =   12
      Top             =   5460
      Width           =   7140
      _Version        =   65536
      _ExtentX        =   12594
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
         Left            =   90
         TabIndex        =   13
         Tag             =   "Select this Code"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   945
         TabIndex        =   14
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1755
         TabIndex        =   15
         Tag             =   "Edit this record"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2580
         TabIndex        =   16
         Tag             =   "Save changes made"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3405
         TabIndex        =   17
         Tag             =   "Cancel changes made"
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4290
         TabIndex        =   18
         Tag             =   "Create a new Record"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5115
         TabIndex        =   19
         Tag             =   "Delete this record"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5940
         TabIndex        =   20
         Tag             =   "Print Listing"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6435
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
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   405
      Left            =   5160
      TabIndex        =   5
      Tag             =   "Find specific record"
      Top             =   4725
      Width           =   720
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
      Left            =   1365
      TabIndex        =   4
      Tag             =   "00-Search Description"
      Top             =   4800
      Width           =   3285
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
      Left            =   120
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "00-Search Code"
      Top             =   4800
      Width           =   1080
   End
   Begin VB.TextBox txtComm 
      Appearance      =   0  'Flat
      DataField       =   "OC_DESCR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Tag             =   "00-Memo field - Free form description"
      Top             =   2925
      Width           =   6615
   End
   Begin VB.TextBox txtSDesc 
      Appearance      =   0  'Flat
      DataField       =   "OC_SDESCR"
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
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   1
      Tag             =   "01-Description of Code"
      Top             =   2550
      Width           =   1695
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "OC_CODE"
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
      Left            =   840
      MaxLength       =   4
      TabIndex        =   0
      Tag             =   "01-NOC Code"
      Top             =   2550
      Width           =   855
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OC_LDATE"
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
      Left            =   120
      MaxLength       =   25
      TabIndex        =   6
      Text            =   "Ldate"
      Top             =   4530
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OC_LTIME"
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
      Left            =   2160
      MaxLength       =   25
      TabIndex        =   7
      Text            =   "LTime"
      Top             =   4530
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OC_LUSER"
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
      Left            =   4200
      MaxLength       =   25
      TabIndex        =   8
      Text            =   "LUser"
      Top             =   4530
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxocclss.frx":0000
      Height          =   2325
      Left            =   120
      OleObjectBlob   =   "fxocclss.frx":0014
      TabIndex        =   9
      Tag             =   "NOC Codes"
      Top             =   60
      Width           =   6870
   End
   Begin VB.Label lblShort 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Abbreviated Description"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   11
      Top             =   2595
      Width           =   2055
   End
   Begin VB.Label lblClass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2595
      Width           =   525
   End
End
Attribute VB_Name = "frmOCCLASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew  As Boolean
Dim rsDATA As New ADODB.Recordset


Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Display_Value


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

glbOClass$ = ""
glbOClassDesc$ = ""
Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()

On Error GoTo DelErr
Dim a%, Msg

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "OCC CLASS", "Delete")
Call RollBack    '11June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "OC_CODE = '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ = "OC_SDESCR>= '" & txtFindDesc.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindDesc = ""
    End If
    Exit Sub
End If

End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()

On Error GoTo Mod_Err
fglbNew = False

Call modSTUPD(True)

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack    '11June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr

Call modSTUPD(True)


fglbNew = True
Call Set_Control("B", Me)
rsDATA.AddNew

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "OCC CLAss", "AddNew")
Resume Next

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim bk
On Error GoTo OK_Err

If Len(txtCode) < 1 Then
    MsgBox "Group is a required field"
    txtCode.SetFocus
    Exit Sub
End If

If Len(txtSDesc) < 1 Then
    MsgBox "Abbrev. Description is a required field"
    txtSDesc.SetFocus
    Exit Sub
End If

If DupNoc Then
    MsgBox "Duplicate NOC Code."
    txtCode.SetFocus
    Exit Sub
End If

Call UpdUStats(Me)

gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, rsDATA)
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh


Call modSTUPD(False)


Exit Sub

OK_Err:
If Err = 3022 Then
    MsgBox "Group " & Left$(txtCode, 4) & " is already in database"
    Exit Sub
    Resume Next
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPROV", "Update")
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

RHeading = "Occupation Classificaitons"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()

glbOClass$ = Data1.Recordset("OC_CODE")
glbOClassDesc$ = Data1.Recordset("OC_SDESCR")
Unload Me

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDEPTS", "SELECT")

End Sub

Private Sub Form_Load()
Dim SQLQ As String
glbOnTop = "FRMOCCLASS"
glbOClass$ = ""
glbOClassDesc$ = ""

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HR_OCCUPATION_CLASS"
Data1.Refresh

'Ticket #19537
If glbCompSerial = "S/N - 2279W" Then
    lblClass.Caption = "Classification"
    txtCode.Left = 1440
    lblShort.Caption = "EEOG"
    vbxTrueGrid.Columns(1).Caption = "EEOG"
Else
    txtCode.Left = 840
End If

Call modSTUPD(False)
If Not gSec_Upd_Job_Classes Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If

Screen.MousePointer = DEFAULT


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

cmdOK.Enabled = TF
cmdCancel.Enabled = TF
cmdNew.Enabled = FT
vbxTrueGrid.Enabled = FT
cmdClose.Enabled = FT
cmdSelect.Enabled = FT
cmdModify.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT
cmdFind.Enabled = FT
txtComm.Enabled = TF
txtCode.Enabled = TF
txtSDesc.Enabled = TF
txtFindKey.Enabled = FT
txtFindDesc.Enabled = FT
If Not glbOClassMode% Or Data1.Recordset.EOF Then
    cmdSelect.Enabled = False
End If
End Sub

Private Sub txtCode_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub

Private Sub txtComm_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtSDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()

If Not Me.vbxTrueGrid.AllowUpdate Then
    glbOClass$ = Data1.Recordset("OC_CODE")
    glbOClassDesc$ = Data1.Recordset("OC_SDESCR")
    If glbOClassMode Then
        Unload Me
    End If
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
    
    SQLQ = "SELECT * FROM HR_OCCUPATION_CLASS"
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    If Me.vbxTrueGrid.AllowUpdate Then
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

''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
       
    SQLQ = "SELECT * FROM HR_OCCUPATION_CLASS where OC_CODE= '" & Data1.Recordset!OC_CODE & "' "
        
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

End Sub

Private Function DupNoc()
    Dim SQLQ
    Dim rsNOC As New ADODB.Recordset
    
    DupNoc = False
    
    SQLQ = "SELECT OC_CODE FROM HR_OCCUPATION_CLASS WHERE OC_CODE= '" & txtCode & "'"
    
    If Not fglbNew Then
        SQLQ = SQLQ & " AND OC_CODE<>'" & Data1.Recordset("OC_CODE") & "'"
    End If
    rsNOC.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsNOC.EOF Then DupNoc = True
    
End Function

