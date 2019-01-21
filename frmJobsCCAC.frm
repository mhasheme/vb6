VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmJobsCCAC 
   Caption         =   "CCAC Positions Lookup"
   ClientHeight    =   5160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin INFOHR_Controls.EmployeeLookup elpEMPNBR 
      DataField       =   "PC_EMPNBR"
      Height          =   285
      Left            =   2220
      TabIndex        =   8
      Top             =   3420
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
      Enabled         =   0   'False
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Tag             =   "Find specific record"
      Top             =   3930
      Width           =   1305
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "PC_POSITION_CONTROL"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      MaxLength       =   6
      TabIndex        =   5
      Tag             =   "01-Position's Code"
      Top             =   3420
      Width           =   1635
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      MaxLength       =   6
      TabIndex        =   4
      Tag             =   "00-Search Code"
      Top             =   3975
      Width           =   1620
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   4500
      Width           =   6645
      _Version        =   65536
      _ExtentX        =   11721
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
         Height          =   375
         Left            =   1785
         TabIndex        =   3
         Tag             =   "Print Departmental Listing"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   945
         TabIndex        =   2
         Tag             =   "Close and exit this screen"
         Top             =   135
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   105
         TabIndex        =   1
         Tag             =   "Select this Department"
         Top             =   135
         Width           =   735
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   4320
         Top             =   480
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6450
         Top             =   150
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmJobsCCAC.frx":0000
      Height          =   3285
      Left            =   0
      OleObjectBlob   =   "frmJobsCCAC.frx":0014
      TabIndex        =   6
      Tag             =   "Skills Lookup"
      Top             =   0
      Width           =   6435
   End
End
Attribute VB_Name = "frmJobsCCAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbPosNbr
Dim fglbPosCode As String
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

SQLQ = "PC_POSITION_CONTROL = '" & txtFindKey.Text & "'"
Data1.Recordset.Requery
Data1.Recordset.Find SQLQ
If Data1.Recordset.EOF Then
    Data1.Refresh
Else
    txtFindKey = ""
End If
txtFindKey.SetFocus
Exit Sub



End Sub

Private Sub cmdFind_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Positions Listing"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdSelect_Click()
'glbJob = Data1.Recordset("JB_CODE")
'glbJobDesc = Data1.Recordset("JB_DESCR")
'Unload frmJOBS
'End Sub

Private Sub cmdSelect_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDEPTS", "SELECT")


End Sub


Private Sub EmployeeLookup1_Change()

End Sub

Private Sub Form_Load()
Dim SQLQ As String

On Error GoTo Job_Err
glbOnTop = "FRMJOBSCCAC"
Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HR_JOB_CONTROL WHERE PC_JOB='" & fglbPosCode & "' ORDER BY PC_POSITION_CONTROL "
Data1.Refresh

Screen.MousePointer = DEFAULT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    cmdSelect.Enabled = False
End If
Call INI_Controls(Me)
txtCode.Enabled = False

Exit Sub
Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Error", "HRJobs", "Cancel")
Resume Next


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub



Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtFindKey_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub vbxTrueGrid_DblClick()
If Not (Data1.Recordset.EOF Or Data1.Recordset.EOF) Then
    frmFind = True
    Call cmdSelect_Click
End If
Unload Me

End Sub

Private Sub vbxTrueGrid_GotFocus()

Call SetPanHelp(ActiveControl)
'txtFindDesc.SetFocus      'Sept 21, Laura
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HR_JOB_CONTROL WHERE PC_JOB='" & fglbPosCode & "' "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    cmdClose.SetFocus
End If

End Sub

Private Sub cmdSelect_Click()
fglbPosNbr = Data1.Recordset("PC_POSITION_CONTROL")
Unload Me
End Sub

Public Property Get PosNbr() As Variant
PosNbr = fglbPosNbr
End Property
Public Property Let PosCode(vData As String)
fglbPosCode = vData
End Property

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
    If IsNull(Data1.Recordset(elpEMPNBR.DataField)) Then
        elpEMPNBR.Text = ""
    Else
        elpEMPNBR.Text = Data1.Recordset(elpEMPNBR.DataField)
    End If
End If
End Sub
