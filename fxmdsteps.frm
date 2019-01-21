VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmMDiscipSteps 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Disciplinary Steps"
   ClientHeight    =   5280
   ClientLeft      =   1350
   ClientTop       =   1650
   ClientWidth     =   8280
   ControlBox      =   0   'False
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
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5280
   ScaleWidth      =   8280
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxmdsteps.frx":0000
      Height          =   3885
      Left            =   120
      OleObjectBlob   =   "fxmdsteps.frx":0014
      TabIndex        =   0
      Tag             =   "Codes Listings"
      Top             =   120
      Width           =   7995
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   4620
      Width           =   8280
      _Version        =   65536
      _ExtentX        =   14605
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
         Appearance      =   0  'Flat
         Caption         =   "&Recalculate"
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Tag             =   "Save the changes made"
         Top             =   150
         Width           =   1695
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   135
         TabIndex        =   5
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   915
         TabIndex        =   6
         Tag             =   "Edit the Information"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1695
         TabIndex        =   7
         Tag             =   "Save the changes made"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2505
         TabIndex        =   8
         Tag             =   "Cancel the changes made"
         Top             =   150
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3360
         TabIndex        =   9
         Tag             =   "Add a new Code"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4170
         TabIndex        =   10
         Tag             =   "Delete code listed above"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   4980
         TabIndex        =   11
         Tag             =   "Print Code Listing Report"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7920
         Top             =   240
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
         Left            =   7320
         Top             =   240
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
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "DS_COMPNO"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6720
      MaxLength       =   3
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "DS_DESC"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   6120
      MaxLength       =   30
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   2025
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "DS_DISCIPLINE"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   13
      Tag             =   "01-Counselling Type- Code"
      Top             =   4080
      Width           =   3915
      _ExtentX        =   6906
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "CETY"
   End
   Begin MSMask.MaskEdBox medStep 
      DataField       =   "DS_STEPNO"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   360
      TabIndex        =   2
      Tag             =   "21-Enter Step Number"
      Top             =   4080
      Width           =   930
      _ExtentX        =   1640
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
      Format          =   "0"
      PromptChar      =   "_"
   End
End
Attribute VB_Name = "frmMDiscipSteps"
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

If Len(medStep) < 1 Then
    MsgBox "Step Number is a required field"
    medStep.SetFocus
    Exit Function
Else
    If Not IsNumeric(medStep) Then
        MsgBox "Step Number must be a integer number"
        medStep.SetFocus
        Exit Function
    End If
End If

If Len(clpCode(0).Text) = 0 Then
    MsgBox "Discipline Code is a required field"
    clpCode(0).SetFocus
    Exit Function
End If

If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
    MsgBox "IF Discipline Code entered it must be known"
    clpCode(0).SetFocus
    Exit Function
End If


If fglbNewRec Then
    'Ky = clpDiv & txtShowKey
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS "
    SQLQ = SQLQ & "WHERE DS_STEPNO = " & Int(Val(medStep)) & " "
    snapTabs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapTabs.BOF And snapTabs.EOF Then
        snapTabs.Close
    Else
        Msg$ = "Step Number already exists in database"
        MsgBox Msg$
        snapTabs.Close
        Exit Function
    End If
Else
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS "
    SQLQ = SQLQ & "WHERE DS_STEPNO = " & Int(Val(medStep)) & " "
    SQLQ = SQLQ & "AND DS_ID <> " & Data1.Recordset("DS_ID")
    snapTabs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If snapTabs.BOF And snapTabs.EOF Then
        snapTabs.Close
    Else
        Msg$ = "Step Number already exists in database"
        MsgBox Msg$
        snapTabs.Close
        Exit Function
    End If
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





Private Sub clpCode_Change(Index As Integer)
'txtDesc = clpCode(0).Caption
End Sub

Private Sub clpCode_LostFocus(Index As Integer)
txtDesc = clpCode(0).Caption
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

Private Sub cmdModify_Click()


On Error GoTo Mod_Err
Call ST_UPD_MODE(True)

medStep.SetFocus

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
txtComp.Text = glbCompNo

medStep.SetFocus
'If clpDiv.Visible Then
'    clpDiv.SetFocus
'Else
'    'txtShowKey.SetFocus
'End If

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

strK = medStep

Data1.Recordset("DS_STEPNO") = medStep
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Data1.Recordset.Find "DS_STEPNO = '" & strK & "'"
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
Me.vbxCrystal.BoundReportHeading = frmMDiscipSteps.Caption
Me.vbxCrystal.WindowTitle = frmMDiscipSteps.Caption & " Report"
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmdRecalc_Click()
If glbWFC And glbPlantCode = "WHBY" Then
    cmdRecalc.Enabled = False
    Call Whitby60daysRule("ALL", "")
    cmdRecalc.Enabled = True
End If
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
glbOnTop = "FRMdISCIPSTEPS"
glbCodeRef = False  'table entrie modified/added false
     

SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS ORDER BY DS_STEPNO"

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = SQLQ
Data1.Refresh

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

medStep.Enabled = TF
txtDesc.Enabled = TF

vbxTrueGrid.Enabled = FT
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If
On Error GoTo ERR_EXIT
'If Not gSec_Upd_Master_Table(glbTabNam) Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'End If
ERR_EXIT:
If Err.Number = 5 Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If
End Sub

Private Sub medStep_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtDesc_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtWaitPeriod_GotFocus()
Call SetPanHelp(ActiveControl)
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
        
        SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub
