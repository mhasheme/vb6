VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEHSAttach 
   AutoRedraw      =   -1  'True
   Caption         =   "Incident Attachments"
   ClientHeight    =   8400
   ClientLeft      =   -135
   ClientTop       =   600
   ClientWidth     =   11400
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DE_TYPE"
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4920
      MaxLength       =   8
      TabIndex        =   20
      Tag             =   "11- Document Number"
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DE_OCCDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6840
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtDocNum 
      Appearance      =   0  'Flat
      DataField       =   "DE_DOCNO"
      Enabled         =   0   'False
      Height          =   285
      Left            =   10320
      MaxLength       =   8
      TabIndex        =   16
      Tag             =   "11- Document Number"
      Top             =   6480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      DataField       =   "DE_DOCDESC"
      Height          =   285
      Left            =   2880
      MaxLength       =   30
      TabIndex        =   15
      Tag             =   "11- Document Number"
      ToolTipText     =   "Click on the lookup icon (magnifying glass) to View the document"
      Top             =   3240
      Width           =   5295
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   325
      Left            =   2880
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fehsatta.frx":0000
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "Fehsatta.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   8055
   End
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   11040
      Top             =   8280
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Caption         =   "Ado3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   11040
      Top             =   8040
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      Caption         =   "Ado1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtShift 
      Appearance      =   0  'Flat
      DataField       =   "DE_CASE"
      Height          =   285
      Left            =   4440
      MaxLength       =   8
      TabIndex        =   2
      Tag             =   "11- incident Number"
      Top             =   2820
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox comShift 
      DataSource      =   "Data1"
      Height          =   315
      Left            =   2880
      TabIndex        =   1
      Tag             =   "01-Incident Number"
      Top             =   2835
      Width           =   1575
   End
   Begin VB.TextBox Updstats 
      DataField       =   "DE_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2670
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7410
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "DE_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4470
      MaxLength       =   25
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7410
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "DE_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6150
      MaxLength       =   25
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   7410
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         Left            =   6480
         TabIndex        =   25
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
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
         TabIndex        =   8
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         AutoSize        =   -1  'True
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
         TabIndex        =   7
         Top             =   135
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   9960
      Top             =   7920
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
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Click on the lookup icon (magnifying glass) to View the document."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   26
      Top             =   4680
      Width           =   6585
   End
   Begin VB.Label lblUserDesc 
      Height          =   255
      Left            =   3840
      TabIndex        =   24
      Top             =   4200
      Width           =   2415
   End
   Begin VB.Label lblUpdateBy 
      Caption         =   "Updated By"
      Height          =   255
      Left            =   2880
      TabIndex        =   23
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label lblUpdDateDesc 
      Height          =   255
      Left            =   7320
      TabIndex        =   22
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Label lblUpdateDate 
      Caption         =   "Updated Date"
      Height          =   255
      Left            =   6240
      TabIndex        =   21
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Document #"
      Height          =   195
      Index           =   6
      Left            =   7800
      TabIndex        =   18
      Top             =   6540
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Document Description"
      Height          =   195
      Index           =   7
      Left            =   360
      TabIndex        =   17
      Top             =   3360
      Width           =   2025
   End
   Begin VB.Image imgNoSec 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2400
      Picture         =   "Fehsatta.frx":2D40
      Top             =   3765
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblImport 
      Alignment       =   1  'Right Justify
      Caption         =   "Incident Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   3780
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image imgSec 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   2400
      Picture         =   "Fehsatta.frx":2E8A
      ToolTipText     =   "Click to View the document"
      Top             =   3765
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblCNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DE_COMPNO"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   720
      TabIndex        =   12
      Top             =   7440
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblEEID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "DE_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1860
      TabIndex        =   11
      Top             =   7440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Incident Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   10
      Top             =   2910
      Width           =   1545
   End
End
Attribute VB_Name = "frmEHSAttach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fGLBNew
Dim rsDATA As New ADODB.Recordset ' Remove Binding Control
Dim oldTarget
Dim oldAssigned
Dim xFieldList1, xFieldList2
Function chkHSAttach()

Dim SQLQ As String, Msg As String, dd#

chkHSAttach = False

On Error GoTo chkHSAttach_Err


Dim tTime As Variant
Dim Part1$, Part2$

Dim RsEHST As New ADODB.Recordset
Dim xDOCNum
'txtDocNum = 1 'Val(Format(txtDocNum, "0000"))
'If Not fglbNew Then
    If Not glbtermopen Then
        SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY WHERE DE_EMPNBR = " & glbLEE_ID
    Else
        SQLQ = "SELECT * FROM TERM_HRDOC_HEALTH_SAFETY WHERE TERM_SEQ = " & glbTERM_Seq
    End If
    If Not fGLBNew Then SQLQ = SQLQ & " AND DE_DOCNO = " & Val(txtDocNum)
    If Not fGLBNew Then
        If Not IsNull(Data1.Recordset("DE_CASE")) Then
            SQLQ = SQLQ & " AND DE_CASE = " & Data1.Recordset("DE_CASE")
        End If
    Else
        SQLQ = SQLQ & " AND DE_CASE = " & Val(txtShift.Text)
    End If
    
    SQLQ = SQLQ & " order by DE_DOCNO desc"
    
    RsEHST.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    If Not RsEHST.EOF Then
        If Not IsNull(RsEHST("DE_DOCNO")) Then
            If fGLBNew Then
            '    txtDocNum = RsEHST("DE_DOCNO")
            'Else
                txtDocNum = RsEHST("DE_DOCNO") + 1
            End If
        End If
    Else
        txtDocNum = 1
    End If
RsEHST.Close
'End If

'~~
If Len(txtShift) < 1 Then
    MsgBox "Incident Number is a required field"
    comShift.SetFocus
    Exit Function
End If
If Not IfIncidentNo(Val(txtShift)) Then
    MsgBox "Incident Number Not Valid"
    comShift.SetFocus
    Exit Function
End If

chkHSAttach = True

Exit Function

chkHSAttach_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSAttach", "HRDOC_HEALTH_SAFETY", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function


Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

fGLBNew = False
Call Display_Value

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE
'Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRDOC_HEALTH_SAFETY", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "frmEHSAttach" Then glbOnTop = ""

End Sub

'Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, X

If Not gSec_Upd_Health_Safety Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err


Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


If glbtermopen Then
    'gdbAdoIhr001_DOC.BeginTrans
    rsDATA.Delete
    'George Feb 16,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_HEALTH_SAFETY where DE_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DE_CASE='" & glbJob & "' and DE_DOCNO='" & txtDocNum & "'" 'and DE_EDATE=" & Date_SQL(glbSDate)
        gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_HEALTH_SAFETY_2 where DE_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DE_CASE='" & glbJob & "' and DE_DOCNO='" & txtDocNum & "'" 'and DE_EDATE=" & Date_SQL(glbSDate)
    End If
    'gdbAdoIhr001_DOC.CommitTrans
    'George Feb 16,2006
    Data1.Refresh
Else
    'gdbAdoIhr001_DOC.BeginTrans
    rsDATA.Delete
    'George fEB 16,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "Delete from HRDOC_HEALTH_SAFETY where DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR = " & glbLEE_ID & " and DE_CASE='" & glbJob & "' and DE_DOCNO='" & txtDocNum & "'" 'and DE_EDATE=" & Date_SQL(glbSDate)
        gdbAdoIhr001_DOC.Execute "Delete from HRDOC_HEALTH_SAFETY_2 where DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR = " & glbLEE_ID & " and DE_CASE='" & glbJob & "' and DE_DOCNO='" & txtDocNum & "'" 'and DE_EDATE=" & Date_SQL(glbSDate)
    End If
    'gdbAdoIhr001_DOC.CommitTrans
    'George Feb 16,2006
    Data1.Refresh
End If
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

Me.vbxTrueGrid.SetFocus
fGLBNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRDOC_HEALTH_SAFETY", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub



Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_Health_Safety Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

fGLBNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
'George on Jan 26,2006 #10266
If gsAttachment_DB Then
    glbJob = ""
    glbSDate = CVDate("01/01/1900")
    'glbDocName = "INCIDENT"
    lblImport.Visible = True
    imgSec.Visible = False
    imgNoSec.Visible = True
    cmdImport.Visible = True
End If
'George on Jan 26,2006 #10266

On Error GoTo AddN_Err
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    Me.vbxTrueGrid.Enabled = False
End If
Me.vbxTrueGrid.Enabled = False
'data1.Recordset.AddNew
Call Set_Control("B", Me)

rsDATA.AddNew '???
fGLBNew = True

oldTarget = ""
oldAssigned = ""
Updstats(4) = glbDocName

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"


Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRDOC_HEALTH_SAFETY", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X, xID
On Error GoTo Add_Err

If Not chkHSAttach() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)


'rsDATA.Requery '???
'If fglbNew Then rsDATA.AddNew

Call Set_Control("U", Me, rsDATA)
If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    'gdbAdoIhr001_DOC.BeginTrans
    rsDATA.Update
    'gdbAdoIhr001_DOC.CommitTrans
    xID = rsDATA("DE_ID")
Else
    'gdbAdoIhr001_DOC.BeginTrans
    Call Set_Control("U", Me, rsDATA)
    rsDATA.Update
    'gdbAdoIhr001_DOC.CommitTrans
    xID = rsDATA("DE_ID")
End If
Data1.Refresh
Data1.Recordset.Find "DE_ID=" & xID

fGLBNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Me.vbxTrueGrid.Enabled = True

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            glbJob = rsDATA("DE_CASE")
            glbDocTmp = rsDATA("DE_DOCNO")
            Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
        End If
    End If
    glbDocImpFile = ""
End If

'Data1.Refresh
'Data1.Recordset.Find "DE_ID=" & xID

'Me.vbxTrueGrid.SetFocus
If NextFormIF("Incident document") Then
    Call cmdNew_Click
End If
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRDOC_HEALTH_SAFETY", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Incident documents"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Incident documents"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
Sub comShift_Change()
'txtShift = comShift  'JDY
End Sub

Sub comShift_Click()
'txtShift = comShift      'JDY
End Sub

Sub comShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comShift_LostFocus()
txtShift = comShift  'JDY
End Sub

Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDOC_HEALTH_SAFETY", "SELECT")


End Sub

Function EERetrieve()
Dim SQLQ As String
EERetrieve = False

Screen.MousePointer = HOURGLASS
On Error GoTo EERError


If glbtermopen Then
    SQLQ = "Select " & xFieldList2 & " from Term_HRDOC_HEALTH_SAFETY"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    'SQLQ = SQLQ & " ORDER BY DE_CASE,DE_DOCNO"
    SQLQ = SQLQ & " ORDER BY DE_CASE DESC,DE_ID DESC"
Else
    SQLQ = "Select " & xFieldList1 & " from HRDOC_HEALTH_SAFETY"
    SQLQ = SQLQ & " where DE_EMPNBR = " & glbLEE_ID
    'SQLQ = SQLQ & " ORDER BY DE_CASE,DE_DOCNO"
    SQLQ = SQLQ & " ORDER BY DE_CASE DESC,DE_ID DESC"
End If

Data1.RecordSource = SQLQ
Data1.Refresh


If glbtermopen Then
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Else
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
End If

Data3.RecordSource = SQLQ
Data3.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function


EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HRDOC_HEALTH_SAFETY", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
Exit Function
End Function
Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "frmEHSAttach"

End Sub

Sub Form_GotFocus()
glbOnTop = "frmEHSAttach"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim X%
Dim SQLQ1
glbOnTop = "frmEHSAttach"
xFieldList1 = "DE_ID,DE_COMPNO,DE_EMPNBR,DE_CASE,DE_OCCDATE,DE_DOCNO,DE_DOCDESC,DE_FILEEXT,DE_TYPE,DE_LDATE,DE_LTIME,DE_LUSER"
xFieldList2 = "DE_ID,DE_COMPNO,DE_EMPNBR,DE_CASE,DE_OCCDATE,DE_DOCNO,DE_DOCDESC,DE_FILEEXT,DE_TYPE,DE_LDATE,DE_LTIME,DE_LUSER,TERM_SEQ"

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRDB_DOC
    Data3.ConnectionString = glbAdoIHRDB
Else
    Data1.ConnectionString = glbAdoIHRDB_DOC
    Data3.ConnectionString = glbAdoIHRDB
End If



Screen.MousePointer = HOURGLASS

If glbLinHS Then 'Ticket #12401
    glbLinEmpNo = glbLEE_ID
    If Not glbtermopen Then
        If Len(glbDiv) = 0 Then Call Get_Div(True) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    Else
        If Len(glbDiv) = 0 Then Call Get_Div(True) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    End If
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
Else
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(Str(glbLEE_ID)), 6) = "999999" Then
                glbLEE_ID = 0
            End If
        End If
    End If
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    Else
        If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
        If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If


If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

comShift.Clear
Do Until Data3.Recordset.EOF                  'JDY
  comShift.AddItem Data3.Recordset("EC_CASE") 'JDY
  Data3.Recordset.MoveNext                    'JDY
Loop

If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Incident documents Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "Incident documents Data - " & Left$(glbLEE_SName, 8)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If

Call ST_UPD_MODE(False)
Call Display_Value

'Updstats(3) = Data3.Recordset("EC_OCCDATE")

If Not gSec_Upd_Health_Safety Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

End Sub

Sub Form_LostFocus()
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

Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSAttach = Nothing 'carmen 18 may 00
Call NextForm
End Sub

Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF
comShift.Enabled = TF

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
Else
'    cmdModify.Enabled = True
End If

'George on Feb 16,2006 #10266
glbJob = ""
glbSDate = "01/01/1900"
glbDocKey = 0
If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
    glbJob = Data1.Recordset("DE_CASE")
    glbSDate = Data1.Recordset("DE_OCCDATE")
    glbDocKey = Data1.Recordset("DE_ID")
    glbDocTmp = Data1.Recordset("DE_DOCNO")
End If
glbDocName = "INCIDENT"
If gsAttachment_DB Then
    Call DispimgIcon(Me, "frmEHSAttach")
    If gSec_Upd_Health_Safety And Not glbtermopen Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If
 
'Updstats(3) = Data3.Recordset("EC_OCCDATE")
Updstats(4) = glbDocName
'George on Feb 16,2006 #10266

End Sub



Sub txtShift_Change()
 
  If Not (Val(txtShift) = 0) Then
    comShift = txtShift
  Else
    comShift = ""
  End If

End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        'If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        'End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
    End If
End Sub

Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbtermopen Then
            SQLQ = "Select " & xFieldList2 & " from Term_HRDOC_HEALTH_SAFETY"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select " & xFieldList1 & " from HRDOC_HEALTH_SAFETY"
            SQLQ = SQLQ & " where DE_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub


''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()

    Dim SQLQ
'    elpAssigned.Caption = ""
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Exit Sub
    End If
'????
If glbtermopen Then
    SQLQ = "Select " & xFieldList2 & " from Term_HRDOC_HEALTH_SAFETY"
    SQLQ = SQLQ & " WHERE DE_ID = " & Data1.Recordset!DE_ID
    SQLQ = SQLQ & " ORDER BY DE_Case DESC"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select " & xFieldList1 & " from HRDOC_HEALTH_SAFETY"
    SQLQ = SQLQ & " where DE_ID = " & Data1.Recordset!DE_ID
    SQLQ = SQLQ & " ORDER BY DE_Case DESC"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic

End If
'???

If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

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
UpdateRight = gSec_Upd_Health_Safety
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
ElseIf rsDATA.EOF Then
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
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Incident documents Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv

    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If

Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSAttach.Caption = "Incident documents Data - " & Left$(glbLEE_SName, 5)
        frmEHSAttach.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    'lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
    
End If
End Sub
Function IfIncidentNo(InciNo As Double)
  IfIncidentNo = False
  If Data3.Recordset.BOF And Data3.Recordset.EOF Then
     Exit Function
  End If
  Data3.Recordset.MoveFirst
  Data3.Recordset.Find "EC_Case=" & InciNo
  If Data3.Recordset.EOF Then Exit Function
  IfIncidentNo = True


End Function


Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmEHSAttach")
    Close   'close all the opened files
    Call FillMemoFile(SQLQ, "INCIDENT")
End Sub

Private Sub cmdImport_Click()
Dim xID
    glbDocNewRecord = fGLBNew
    glbDocName = "INCIDENT"
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        glbDocKey = 0
        glbJob = ""
        glbDocTmp = ""
    Else
        glbDocKey = Data1.Recordset("DE_ID")
        glbJob = rsDATA("DE_CASE")
        glbDocTmp = rsDATA("DE_DOCNO")
    End If

    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEHSAttach")
    
'    If Not glbDocNewRecord Then
'        xID = rsDATA("DE_ID")
'        Data1.Refresh
'        Data1.Recordset.Find "DE_ID=" & xID
'    End If
End Sub

