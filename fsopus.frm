VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmOpus 
   Appearance      =   0  'Flat
   Caption         =   "Intellisol Matrix"
   ClientHeight    =   5565
   ClientLeft      =   105
   ClientTop       =   1395
   ClientWidth     =   10020
   FillStyle       =   0  'Solid
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5565
   ScaleWidth      =   10020
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "INFOHR Code"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   3120
      TabIndex        =   3
      Top             =   4230
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6360
      Top             =   5160
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      TabIndex        =   8
      Top             =   4905
      Width           =   10020
      _Version        =   65536
      _ExtentX        =   17674
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
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   210
         TabIndex        =   9
         Tag             =   "Close and exit this screen"
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1050
         TabIndex        =   10
         Tag             =   "Edit the information on this screen"
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1890
         TabIndex        =   11
         Tag             =   "Save the changes made"
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2730
         TabIndex        =   12
         Tag             =   "Cancel changes made"
         Top             =   150
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   3735
         TabIndex        =   13
         Tag             =   "Print Listing "
         Top             =   150
         Visible         =   0   'False
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   5730
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
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.TextBox txtCompPaid 
      Appearance      =   0  'Flat
      DataField       =   "Percent Company Paid"
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
      Left            =   8760
      TabIndex        =   6
      Tag             =   "11-Enter % of Company Paid"
      Top             =   4215
      Width           =   1095
   End
   Begin VB.TextBox txtEmpPaid 
      Appearance      =   0  'Flat
      DataField       =   "Percent Employee Paid"
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
      Left            =   7040
      TabIndex        =   5
      Tag             =   "11-Enter % of Employee Paid"
      Top             =   4215
      Width           =   1095
   End
   Begin VB.TextBox txtACodeINFO 
      Appearance      =   0  'Flat
      DataField       =   "INFOHR Code Table Name"
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
      Left            =   1920
      TabIndex        =   2
      Tag             =   "01-INFO HR Code"
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox txtPayrcode 
      Appearance      =   0  'Flat
      DataField       =   "Payrcode"
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
      Left            =   240
      TabIndex        =   1
      Tag             =   "01-Payrol Code"
      Top             =   4200
      Width           =   1095
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fsopus.frx":0000
      Height          =   3555
      Left            =   240
      OleObjectBlob   =   "fsopus.frx":0014
      TabIndex        =   0
      Tag             =   "List of Codes"
      Top             =   360
      Width           =   9585
   End
   Begin MSMask.MaskEdBox medTaxBen 
      DataField       =   "Taxable Benefit"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5640
      TabIndex        =   4
      Tag             =   "31-Enter Yes /No"
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
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
      Format          =   "yes/no"
      PromptChar      =   "_"
   End
   Begin VB.Label lblCodeDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Unassigned"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   3660
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   840
   End
End
Attribute VB_Name = "frmOpus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CurRec As Variant
Dim vbxComp As String
Dim vbxEmp As String
Dim fglbNew As Boolean


Private Function chkEntry()
Dim oCode As String, OCodeD As String

chkEntry = False

'If Len(clpcode(1)) = 0 Then
'    MsgBox "The User Code is a required field"
'    clpcode(1).Enabled = True
'    clpcode(1).SetFocus
'    Exit Function
'End If

'If lblCodeDesc(1) = "Unassigned" And Len(clpCode(1)) <> 0 Then
'    MsgBox "The User Code must be valid"
'    Exit Function
'End If

If UCase(medTaxBen) = "YES" Or UCase(medTaxBen) = "TRUE" Then
    medTaxBen = -1
    ElseIf Val(medTaxBen) = -1 Then
        medTaxBen = -1
    ElseIf UCase(medTaxBen) = "NO" Or UCase(medTaxBen) = "FALSE" Then
        medTaxBen = 0
    ElseIf Val(medTaxBen) = 0 Then
        medTaxBen = 0
Else
    MsgBox "You must enter Yes / No!"
    Exit Function
End If

If Not IsNumeric(txtEmpPaid) Then
    MsgBox "The value must be numeric!"
    txtEmpPaid = ""
    txtEmpPaid.SetFocus
    Exit Function
End If

If Not IsNumeric(txtCompPaid) Then
    MsgBox "The value must be numeric!"
    txtCompPaid = ""
    txtCompPaid.SetFocus
    Exit Function
End If

If (Val(txtEmpPaid) + Val(txtCompPaid)) <> 100 And (Val(txtEmpPaid) + Val(txtCompPaid)) <> 0 Then
    MsgBox "The sum between Percent Employee Paid and Percent Company Paid must be 100 or 0!"
    txtCompPaid.SetFocus
    Exit Function
End If

chkEntry = True

End Function

Private Sub clpCode1_Change()

End Sub

Private Sub clpCode_Change(Index As Integer)
lblCodeDesc(Index) = clpCode(Index).Caption
End Sub

Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

'Data1.UpdateControls    ' returns without saving
bk = Data1.Recordset.Bookmark
Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Data1.Recordset.Bookmark = bk

'vbxTrueGrid.Enabled = True
'cmdOK.Enabled = False
'cmdCancel.Enabled = False
'cmdPrint.Enabled = True
'cmdClose.Enabled = True
'cmdEdit.Enabled = True
medTaxBen.Enabled = False
'clpCode(1).Enabled = False
txtEmpPaid.Enabled = False
txtCompPaid.Enabled = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "IHROPUS", "Cancel")
Call RollBack '08June99 js

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdClose_Click()

Unload Me
glbOnTop = "FRMEREINSTATE"

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdEdit_Click()

On Error GoTo Mod_Err

'vbxTrueGrid.Enabled = False
'cmdPrint.Enabled = False
'cmdClose.Enabled = False
'cmdEdit.Enabled = False
'cmdOK.Enabled = True
'cmdCancel.Enabled = True
medTaxBen.Enabled = True
clpCode(1).Enabled = True
clpCode(1).SetFocus
If Data1.Recordset("INFOHR Code Table Name") = "BNCD" Then
    txtCompPaid.Enabled = True
    txtEmpPaid.Enabled = True
Else
    txtCompPaid.Enabled = False
    txtEmpPaid.Enabled = False
End If
Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRMATRIX", "Modify")
Call RollBack '08June99 js

End Sub

Private Sub cmdEdit_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdOK_Click()
Dim bk, strT
On Error GoTo Add_Err

If Not chkEntry() Then Exit Sub
bk = Data1.Recordset.Bookmark
Data1.Recordset("Payrcode") = txtPayrcode & ""
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Data1.Recordset.Bookmark = bk

Me.vbxTrueGrid.Enabled = True

'cmdCancel.Enabled = False
'cmdOK.Enabled = False
'cmdPrint.Enabled = True
'cmdEdit.Enabled = True
'cmdClose.Enabled = True
medTaxBen.Enabled = False
'clpCode(1).Enabled = False
txtEmpPaid.Enabled = False
txtCompPaid.Enabled = False

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdOK", "IHROPUS", "Update")
Call RollBack '08June99 js

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = Me.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub
Sub cmdView_Click()
Dim RHeading As String

RHeading = Me.Caption
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM paycode_infohr ORDER BY Payrcode"
Data1.RecordSource = SQLQ
Data1.Refresh
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Attendance", "HR Attendance", "SELECT")
Call RollBack '08June99 js

Exit Function

End Function

Private Sub Form_Activate()
    glbOnTop = "FRMINFOATTEND"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMINFOATTEND"
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

On Error Resume Next
Dim x% ' records found
Dim SQLQ As String
Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM paycode_infohr ORDER BY Payrcode"

If glbSQL Then
    Data1.ConnectionString = glbAdoIHRDB
Else
    Data1.ConnectionString = glbAdoIHRDBO
End If

Data1.RecordSource = SQLQ
Data1.Refresh

If Data1.Recordset.RecordCount <= 0 Then
  cmdEdit.Enabled = False
End If

Screen.MousePointer = DEFAULT

If gSec_Matrix Then             'May99 js
    cmdCancel.Enabled = False    '
    cmdOK.Enabled = False        '
    'medTaxBen.Enabled = False    '
    txtACodeINFO.Enabled = False '
    'clpCode(1).Enabled = False   '
    txtCompPaid.Enabled = False  '
    txtEmpPaid.Enabled = False   '
    txtPayrcode.Enabled = False  '
End If                           '
    Call INI_Controls(Me)
End Sub

Private Sub medTaxBen_GotFocus()
'    medTaxBen = ""
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtACodeINFO_Change()
    clpCode(1).TablName = txtACodeINFO.Text
End Sub

Private Sub txtACodeINFO_GotFocus()
    Call SetPanHelp(ActiveControl)
    clpCode(1).TablName = txtACodeINFO
End Sub

Private Sub txtACodeINFO_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

End Sub




Private Sub txtCompPaid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEmpPaid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPayrcode_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtPayrcode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))

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
        
        SQLQ = "SELECT * FROM paycode_infohr"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    vbxComp = vbxComp + (Chr$(KeyAscii))
End Sub

Private Sub vbxTrueGrid_Update(Row As Long, Col As Integer, Value As String)

If (Val(txtCompPaid) + Val(txtEmpPaid)) > 100 Then
    'Data1.Recordset.Edit
    Data1.Recordset("Percent Company Paid") = 0
    'Data1.UpdateRecord
    Data1.Recordset.UpdateBatch
    Data1.Recordset.Resync
    
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
UpdateRight = True
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
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
    UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub
