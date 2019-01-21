VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmECOMMENTS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Comments"
   ClientHeight    =   9645
   ClientLeft      =   105
   ClientTop       =   1035
   ClientWidth     =   14355
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
   ScaleHeight     =   9645
   ScaleWidth      =   14355
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Frank Test Get Value"
      Height          =   495
      Left            =   9360
      TabIndex        =   22
      Top             =   8040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Frank Test Add"
      Height          =   495
      Left            =   9360
      TabIndex        =   21
      Top             =   7440
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   8220
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fecommnt.frx":0000
      Height          =   1845
      Left            =   180
      OleObjectBlob   =   "fecommnt.frx":0014
      TabIndex        =   0
      Top             =   660
      Width           =   8895
   End
   Begin INFOHR_Controls.DateLookup dlpEDate 
      DataField       =   "CO_EDATE"
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Tag             =   "41-Effective Date of Comment"
      Top             =   3000
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "CO_TYPE"
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Tag             =   "01-Comment Type- Code"
      Top             =   2640
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ECOM"
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   15
      Top             =   8985
      Width           =   14355
      _Version        =   65536
      _ExtentX        =   25321
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
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "CO_COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "00-Comments"
      Top             =   3690
      Width           =   8805
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CO_LDATE"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CO_LTIME"
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
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CO_LUSER"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   14355
      _Version        =   65536
      _ExtentX        =   25321
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
         TabIndex        =   19
         Top             =   115
         Width           =   1305
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
         TabIndex        =   18
         Top             =   115
         Width           =   720
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   110
         Width           =   1245
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1335
      Left            =   7920
      TabIndex        =   20
      Top             =   5880
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2355
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Select"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Question"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.Image imgNoSec 
      Height          =   240
      Left            =   7800
      Picture         =   "fecommnt.frx":2EE4
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblImport 
      Alignment       =   1  'Right Justify
      Caption         =   "Comments"
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
      Height          =   255
      Left            =   5880
      TabIndex        =   17
      Top             =   2640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Image imgSec 
      Height          =   240
      Left            =   7800
      Picture         =   "fecommnt.frx":302E
      Top             =   2640
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   14
      Top             =   3030
      Width           =   780
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   13
      Top             =   2670
      Width           =   435
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Height          =   255
      Index           =   3
      Left            =   330
      TabIndex        =   12
      Top             =   3390
      Width           =   2910
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CO_EMPNBR"
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
      TabIndex        =   10
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
      DataField       =   "CO_COMPNO"
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
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmECOMMENTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddChg
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbAddable

Private Function chkEComment()
Dim SQLQ As String, Msg As String, dd#
Dim rs As New ADODB.Recordset
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)

chkEComment = False

On Error GoTo chkEComment_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Comment Type is a required field"
    clpCode(1).SetFocus
    Exit Function
End If
 
If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Comment Type must be valid"
    clpCode(1).SetFocus
    Exit Function
Else
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_COMMENTS WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_COMMENTS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    
    SQLQ = SQLQ & " AND CODENAME='" & clpCode(1).Text & "'"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If rs("MAINTAINABLE") = 0 Then
            MsgBox "You do not have Authority for this Transaction", vbOKOnly + vbInformation, "Authorization failed"
            Exit Function
        End If
    Else
        MsgBox "You do not have Authority for this Transaction", vbOKOnly + vbInformation, "Authorization failed"
        Exit Function
    End If
End If

If Len(dlpEDATE.Text) >= 1 Then
    If Not IsDate(dlpEDATE.Text) Then
        MsgBox "Effective Date is not a valid date."
        dlpEDATE.SetFocus
        Exit Function
    End If
Else
    MsgBox "Effective Date is required."
    dlpEDATE.SetFocus
    Exit Function
End If

If Len(memComments) > 4000 Then 'Ticket #10899
Dim Mag, a%
    Msg = "The Comments field can only contain 4000 charaters. " & Chr(10)
    Msg = Msg & "You have typed more than 4000 charaters." & Chr(10)
    Msg = Msg & "Are you sure you still want to save it?"
    a% = MsgBox(Msg, vbYesNo + vbQuestion, "Confirm Save")
    If a% = vbNo Then
        Exit Function
    End If
    memComments = Left(memComments, 4000)
End If

chkEComment = True

Exit Function

chkEComment_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkComments", "HR_COMMENTS", "edit/Add")
Call RollBack '28July99 js

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

'Data1.UpdateControls    ' returns without saving
'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
''' Sam add July 2002 * Remove Binding Control

fglbNew = False
Call SET_UP_MODE

rsDATA.CancelUpdate
Call Display_Value




'Call ST_UPD_MODE(True)  ' reset screen's attributes

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_COMMENTS", "Cancel")
Call RollBack '28July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()

Unload Me
If glbOnTop = "FRMECOMMENTS" Then glbOnTop = ""
Call NextForm
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x
 
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

If Not Comment_Sec Then
    MsgBox "You do not have Authority to complete this transaction.", vbInformation + vbOKOnly, "Authorization failure"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, vbYesNo + vbQuestion, "Confirm Delete")

If a% = vbNo Then Exit Sub


If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    'George Jan 26,2006
    If gsAttachment_DB Then
        'gdbAdoIhr001_DOC.BeginTrans
        'gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_COMMENTS where DO_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DO_COTYPE='" & glbJob & "' and DO_EDATE=" & Date_SQL(glbSDate)
        gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_COMMENTS where DO_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and DO_DOCKEY=" & glbDocKey & " "
        'gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
    Data1.Refresh
Else
    'gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    'gdbAdoIhr001.CommitTrans
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Delete from HRDOC_COMMENTS where DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR = " & glbLEE_ID & " and DO_DOCKEY=" & glbDocKey & " " '
        gdbAdoIhr001_DOC.CommitTrans
    End If
    
    'Ticket #18806
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    'George Jan 26,2006
    Data1.Refresh
End If

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Call Display_Value
End If

 fglbNew = False

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_COMMENTS", "Delete")
Call RollBack '28July99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err
Call SET_UP_MODE
'Call ST_UPD_MODE(True)
'clpCode(1).Enabled = True
'clpCode(1).SetFocus
AddChg = "C"
fglbNew = False
Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_COMMENTS", "Modify")
Call RollBack '28July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

'George on Jan 26,2006 #10266
If gsAttachment_DB Then
    glbJob = ""
    glbSDate = "01/01/1900"
    lblImport.Visible = True 'False
    imgSec.Visible = False
    imgNoSec.Visible = True 'False
    cmdImport.Visible = True 'False
End If
'George on Jan 26,2006 #10266

clpCode(1).Enabled = True
clpCode(1).SetFocus

On Error GoTo AddN_Err


Call Set_Control("B", Me)
rsDATA.AddNew


If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

glbCommentType = ""
glbCommentDate = ""

AddChg = "A"
fglbNew = True
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_COMMENTS", "Add")
Call RollBack '28July99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x, xID
Dim rsCOM As New ADODB.Recordset

On Error GoTo Add_Err

'Release 8.1
If (gSec_Add_Comments And Not gSec_Upd_Comments) And AddChg = "C" Then
    MsgBox "You do not have authority to make changes"
    Call cmdCancel_Click
    Exit Sub
End If

If Not chkEComment() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    rsDATA.Resync
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_COMMENTS set DO_COTYPE='" & rsDATA("CO_TYPE") & "',DO_EDATE=" & Date_SQL(rsDATA("CO_EDATE")) & " where DO_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " AND DO_DOCKEY = " & glbDocKey '& " and DO_EDATE=" & Date_SQL(glbSDate)
        gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
    xID = rsDATA!CO_COMMENT_ID
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    rsDATA.Resync
    'George Jan 26,2006
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.BeginTrans
        gdbAdoIhr001_DOC.Execute "Update HRDOC_COMMENTS set DO_COTYPE='" & rsDATA("CO_TYPE") & "',DO_EDATE=" & Date_SQL(rsDATA("CO_EDATE")) & " where DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR = " & glbLEE_ID & " AND DO_DOCKEY = " & glbDocKey '& " and DO_EDATE=" & Date_SQL(glbSDate)
        gdbAdoIhr001_DOC.CommitTrans
    End If
    'George Jan 26,2006
    xID = rsDATA!CO_COMMENT_ID
End If
Data1.Refresh

Data1.Recordset.Find "CO_COMMENT_ID=" & xID
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Me.vbxTrueGrid.SetFocus

If gsAttachment_DB Then
    If glbDocNewRecord Then 'New Record only
        If Len(glbDocImpFile) > 0 Then
            glbDocKey = xID
            If glbtermopen Then
                Call AttachmentAdd(glbTERM_ID, glbDocImpFile, glbDocType, glbDocDesc)
            Else
                Call AttachmentAdd(glbLEE_ID, glbDocImpFile, glbDocType, glbDocDesc)
            End If
        End If
    End If
    glbDocImpFile = ""
End If

If NextFormIF("Comment") Then
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_COMMENTS", "Update")
Call RollBack '28July99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String, dscGroup$

    RHeading = lblEEName & "'s " & lStr("Comments")
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    Me.vbxCrystal.Formulas(1) = "lblComDesc = '" & lStr("Comments") & "'"
    
    If Not glbtermopen Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgcomme.rpt"
        Me.vbxCrystal.SelectionFormula = "{HR_COMMENTS.CO_EMPNBR} = " & glbLEE_ID
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRDB
            Me.vbxCrystal.DataFiles(1) = glbIHRDB
        End If
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgcommet.rpt"
        Me.vbxCrystal.SelectionFormula = "{Term_COMMENTS.TERM_SEQ} = " & glbTERM_Seq
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
            Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
        End If
    End If
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
    
End Sub

Sub cmdView_Click()
Dim RHeading As String, dscGroup$

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    RHeading = lblEEName & "'s Comments"
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    Me.vbxCrystal.Formulas(1) = "lblComDesc = '" & lStr("Comments") & "'"
    
    If Not glbtermopen Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgcomme.rpt"
        Me.vbxCrystal.SelectionFormula = "{HR_COMMENTS.CO_EMPNBR} = " & glbLEE_ID
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRDB
            Me.vbxCrystal.DataFiles(1) = glbIHRDB
        End If
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgcommet.rpt"
        Me.vbxCrystal.SelectionFormula = "{Term_COMMENTS.TERM_SEQ} = " & glbTERM_Seq
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
            Me.vbxCrystal.DataFiles(1) = glbIHRAUDIT
        End If
    End If
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1
   ' cmdPrint.Enabled = True
    
End Sub


'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Function EERetrieve()
Dim SQLQ As String
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_Comments_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("Comments") & ".", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If

If glbtermopen Then         'Lucy July 5, 2000
     'Added by bryan 14/Oct/05 Ticket#9424
    'call buildSec
    SQLQ = "Select * from Term_COMMENTS "
    SQLQ = SQLQ & " WHERE Term_COMMENTS.TERM_SEQ = " & glbTERM_Seq & " AND Term_COMMENTS.CO_TYPE " & buildSec
Else
    'Added by bryan 12/Oct/05 Ticket#9424
    If glbOracle Then
        SQLQ = "Select HR_COMMENTS.* from HR_COMMENTS, HR_SECURE_COMMENTS"
        SQLQ = SQLQ & " where HR_COMMENTS.CO_TYPE = HR_SECURE_COMMENTS.CODENAME AND CO_EMPNBR = " & glbLEE_ID
        If xTemplate = "" Or xTemplate = "TEMPLATE" Then
            SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
        Else
            '????Ticket #24808 -  Retrieve template's security profile
            SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
        End If
    Else
        SQLQ = "Select HR_COMMENTS.* from HR_COMMENTS INNER JOIN HR_SECURE_COMMENTS ON HR_COMMENTS.CO_TYPE = HR_SECURE_COMMENTS.CODENAME"
        SQLQ = SQLQ & " where CO_EMPNBR = " & glbLEE_ID
        If xTemplate = "" Or xTemplate = "TEMPLATE" Then
            SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
        Else
            '????Ticket #24808 -  Retrieve template's security profile
            SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
        End If
    End If
End If
SQLQ = SQLQ & " ORDER BY CO_EDATE DESC,CO_TYPE"
Data1.RecordSource = SQLQ
Data1.Refresh

'glbJob = "" 'George on Jan 24,2006 #10266
'glbSDate = "01/01/1900" 'George on Jan 24,2006 #10266
'If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
'    glbJob = Data1.Recordset("CO_TYPE") 'George on Jan 24,2006 #10266
'    glbSDate = Data1.Recordset("CO_EDATE") 'George on Jan 24,2006 #10266
'End If
'glbDocName = "Comments"
'If gsAttachment_DB Then
'    Call DispimgIcon(Me, "frmECOMMENTS")
'End If

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SklsRetrieve", "HR_COMMENTS", "SELECT")
Call RollBack '28July99 js

Exit Function

End Function

Private Sub Command1_Click()
        lvw.ColumnHeaders(2).Text = "Division Description"
        lvw.ColumnHeaders(3).Text = "Division Code"
        lvw.ListItems.Add
        lvw.ListItems(lvw.ListItems.count).Checked = True
        'lvw.ListItems(lvw.ListItems.count).Checked = False
        lvw.ListItems(lvw.ListItems.count).SubItems(1) = "Question " & lvw.ListItems.count
        lvw.ListItems(lvw.ListItems.count).SubItems(2) = "a"
        lvw.ListItems.Add
        lvw.ListItems(lvw.ListItems.count).Checked = False
        lvw.ListItems(lvw.ListItems.count).SubItems(1) = "Question " & lvw.ListItems.count
        lvw.ListItems(lvw.ListItems.count).SubItems(2) = "b"
        
        lvw.ListItems.Add
        lvw.ListItems(lvw.ListItems.count).Checked = False
        lvw.ListItems(lvw.ListItems.count).SubItems(1) = "Question " & lvw.ListItems.count
        lvw.ListItems(lvw.ListItems.count).SubItems(2) = "c"
        
        lvw.ListItems.Add
        lvw.ListItems(lvw.ListItems.count).Checked = False
        lvw.ListItems(lvw.ListItems.count).SubItems(1) = "Question " & lvw.ListItems.count
        lvw.ListItems(lvw.ListItems.count).SubItems(2) = "d"
        
        lvw.ListItems.Add
        lvw.ListItems(lvw.ListItems.count).Checked = False
        lvw.ListItems(lvw.ListItems.count).SubItems(1) = "Question " & lvw.ListItems.count
        lvw.ListItems(lvw.ListItems.count).SubItems(2) = "e"
        
        lvw.ListItems.Add
        lvw.ListItems(lvw.ListItems.count).Checked = False
        lvw.ListItems(lvw.ListItems.count).SubItems(1) = "Question " & lvw.ListItems.count
        lvw.ListItems(lvw.ListItems.count).SubItems(2) = "f"
        
End Sub

Private Sub Command2_Click()
Dim xTot As Integer
Dim I As Integer
Dim xList As String
    xTot = lvw.ListItems.count
    
    xList = ""
    For I = 1 To xTot
        If lvw.ListItems(I).Checked Then
            xList = xList & "'" & lvw.ListItems(I).SubItems(2) & "',"
        End If
    Next
    
    MsgBox xList
    'lvw.ListItems(lvw.ListItems.count).Checked = True
    ''lvw.ListItems(lvw.ListItems.count).Checked = False
    'lvw.ListItems(lvw.ListItems.count).SubItems(1) = "Question " & lvw.ListItems.count
    'lvw.ListItems(lvw.ListItems.count).SubItems(2) = "a"
        
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
'Me.cmdModify_Click
    glbOnTop = "FRMECOMMENTS"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMECOMMENTS"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMECOMMENTS"
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

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_Comments_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("Comments") & ".", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Sub
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

Me.vbxTrueGrid.SetFocus
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Comments - " & Left$(glbLEE_SName, 5)
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

If gSec_Upd_Comments Then
    Call ST_UPD_MODE(True)
'Else
'    Call ST_UPD_MODE(False)             '
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

'Comments labels
frmECOMMENTS.Caption = lStr(frmECOMMENTS.Caption)
lblTitle(3).Caption = lStr(lblTitle(3).Caption)
lblImport.Caption = lStr(lblImport.Caption)
vbxTrueGrid.Columns(2).Caption = lStr(vbxTrueGrid.Columns(2).Caption)

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
    Set frmECOMMENTS = Nothing 'carmen may 00
    Call NextForm
End Sub

Private Sub memComments_GotFocus()

Call SetPanHelp(ActiveControl)
MDIMain.panHelp(2).Caption = " "

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
memComments.Locked = FT
clpCode(1).Enabled = TF
dlpEDATE.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'vbxTrueGrid.Enabled = FT


glbDocName = "Comments"
If gsAttachment_DB Then
    'George on Jan 26,2006 #10266
    'glbJob = ""
    'glbSDate = "01/01/1900" 'George on Jan 24,2006 #10266
    If Not (rsDATA.BOF And rsDATA.EOF) Then
        'glbJob = rsDATA("CO_TYPE") 'George on Jan 24,2006 #10266
        'glbSDate = rsDATA("CO_EDATE") 'George on Jan 24,2006 #10266
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("CO_DOCKEY")) Then
                glbDocKey = rsDATA("CO_DOCKEY")
            Else
                glbDocKey = 0
            End If
        Else
            If Not Data1.Recordset.EOF Then
                If Not IsNull(Data1.Recordset("CO_DOCKEY")) Then
                    glbDocKey = Data1.Recordset("CO_DOCKEY")
                Else
                    glbDocKey = 0
                End If
            Else
                glbDocKey = 0
            End If
        End If
    End If
    Call DispimgIcon(Me, "frmECOMMENTS")
    If gSec_Upd_Comments Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False 'George on Jan 26,2006 #10266
        Else
            cmdImport.Visible = True 'George on Jan 26,2006 #10266
        End If
    End If
End If
'George on Jan 26,2006 #10266

fUPMode = TF    ' update mode

End Sub


Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(glbUserID)
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbtermopen Then         'Lucy July 5, 2000
            SQLQ = "Select * from Term_COMMENTS "
            SQLQ = SQLQ & " WHERE Term_COMMENTS.TERM_SEQ = " & glbTERM_Seq & " AND Term_COMMENTS.CO_TYPE " & buildSec
        Else
            'Added by bryan 12/Oct/05 Ticket#9424
            If glbOracle Then
                SQLQ = "Select HR_COMMENTS.* from HR_COMMENTS, HR_SECURE_COMMENTS"
                SQLQ = SQLQ & " where HR_COMMENTS.CO_TYPE = HR_SECURE_COMMENTS.CODENAME AND CO_EMPNBR = " & glbLEE_ID
                If xTemplate = "" Or xTemplate = "TEMPLATE" Then
                    SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
                Else
                    '????Ticket #24808 -  Retrieve template's security profile
                    SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
                End If
            Else
                SQLQ = "Select HR_COMMENTS.* from HR_COMMENTS INNER JOIN HR_SECURE_COMMENTS ON HR_COMMENTS.CO_TYPE = HR_SECURE_COMMENTS.CODENAME"
                SQLQ = SQLQ & " where CO_EMPNBR = " & glbLEE_ID
                If xTemplate = "" Or xTemplate = "TEMPLATE" Then
                    SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
                Else
                    '????Ticket #24808 -  Retrieve template's security profile
                    SQLQ = SQLQ & " AND HR_SECURE_COMMENTS.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_COMMENTS.ACCESSABLE <> 0"
                End If
            End If
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_COMMENTS", "Add")
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
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            'Changed by Bryan 14/Oct/05 Ticket#9424
            'rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            rsDATA.Open "SELECT * FROM Term_COMMENTS WHERE TERM_SEQ=" & glbTERM_Seq, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open "Select * from HR_COMMENTS", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Me.cmdModify_Click
       Exit Sub
    End If
    
    
If glbtermopen Then
    SQLQ = "Select * from Term_COMMENTS"
    SQLQ = SQLQ & " WHERE CO_COMMENT_ID = " & Data1.Recordset!CO_COMMENT_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HR_COMMENTS"
    SQLQ = SQLQ & " where CO_COMMENT_ID = " & Data1.Recordset!CO_COMMENT_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
SQLQ = SQLQ & " ORDER BY CO_EDATE DESC,CO_TYPE"

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE
Me.cmdModify_Click
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
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
'Release 8.1
'UpdateRight = gSec_Upd_Comments
UpdateRight = fglbAddable Or gSec_Upd_Comments

End Property

Public Property Get Addable() As Boolean
'Release 8.1
'Addable = True
Addable = fglbAddable

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
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = Comment_Sec()
End If

'Release 8.1
fglbAddable = gSec_Add_Comments

Call set_Buttons(UpdateState)

If Not UpdateRight Then TF = False

Call ST_UPD_MODE(TF)

End Sub

Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmECOMMENTS.Caption = lStr("Comments - ") & Left$(glbLEE_SName, 5)
    frmECOMMENTS.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Function Comment_Sec() As Boolean
    'created by Bryan 12/Oct/05 Ticket#9424
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As Boolean
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
    
    strSQL = "SELECT MAINTAINABLE from HR_SECURE_COMMENTS WHERE "
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        retVal = Abs(rs("MAINTAINABLE"))
    Else
        retVal = False
    End If
    
    Comment_Sec = retVal
End Function

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmECOMMENTS")
    Call FillMemoFile(SQLQ, "Comments")
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "Comments"
    
    If fglbNew Then
        glbDocKey = 0
    Else
        'Ticket #28839
        'glbDocKey = rsDATA("CO_COMMENT_ID") 'Ticket #16018
        If Not IsNull(rsDATA("CO_DOCKEY")) Then
            glbDocKey = rsDATA("CO_DOCKEY")
        Else
            glbDocKey = rsDATA("CO_COMMENT_ID") 'Ticket #16018
        End If
        glbCommentType = rsDATA("CO_TYPE")
        glbCommentDate = rsDATA("CO_EDATE")
    End If
    
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmECOMMENTS")
End Sub

