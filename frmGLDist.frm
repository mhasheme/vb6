VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmEGLDist 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "G/L Distribution"
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmGLDist.frx":0000
      Height          =   1845
      Left            =   180
      OleObjectBlob   =   "frmGLDist.frx":0014
      TabIndex        =   0
      Top             =   660
      Width           =   8895
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   16
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
      Begin VB.CommandButton cmdCalcPer 
         Caption         =   "Verify %"
         Height          =   495
         Left            =   360
         TabIndex        =   19
         Top             =   0
         Width           =   1335
      End
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
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "GL_COMMENTS"
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
      Height          =   1425
      Left            =   2250
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Tag             =   "00-Comments"
      Top             =   4080
      Width           =   5115
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "GL_LDATE"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "GL_LTIME"
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "GL_LUSER"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
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
         Left            =   7320
         TabIndex        =   21
         Top             =   120
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   115
         Width           =   720
      End
   End
   Begin INFOHR_Controls.CodeLookup clpGLNbr 
      DataField       =   "GL_GLNO"
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Tag             =   "00-GL #"
      Top             =   3000
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   3
   End
   Begin MSMask.MaskEdBox medPPct 
      DataField       =   "GL_PERCENT"
      Height          =   285
      Left            =   2250
      TabIndex        =   3
      Tag             =   "10-Enter Pay Percentage "
      Top             =   3720
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.DateLookup dlpEDate 
      DataField       =   "GL_EDATE"
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "41-Effective Date"
      Top             =   3360
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpOHRSDept 
      DataField       =   "GL_OHRSDEPTNO"
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Tag             =   "00-OHRS Department"
      Top             =   5640
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   14
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "OHRS Department #"
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
      TabIndex        =   22
      Top             =   5685
      Width           =   1485
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   20
      Top             =   3360
      Width           =   1245
   End
   Begin VB.Label lblExp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
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
      Top             =   3720
      Width           =   825
   End
   Begin VB.Label lblGL 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "G/L #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   330
      TabIndex        =   17
      Top             =   3030
      Width           =   525
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
      TabIndex        =   15
      Top             =   4050
      Width           =   990
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "GL_EMPNBR"
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
      TabIndex        =   13
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
      DataField       =   "GL_COMPNO"
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
      TabIndex        =   14
      Top             =   6000
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEGLDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Actn
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset
Dim fglbNew
Dim oGL, oPCT, oDATE, oOHRSDept

Private Function chkGLDist()
Dim SQLQ As String, Msg As String, dd#

chkGLDist = False

On Error GoTo chkGLDist_Err

If Len(clpGLNbr.Text) = 0 Then
    MsgBox lStr("GL # is required")
    clpGLNbr.SetFocus
    Exit Function
End If

If clpGLNbr.Caption = "Unassigned" Then
    MsgBox lStr("GL # must be valid")
    clpGLNbr.SetFocus
    Exit Function
End If

If Len(dlpEDate.Text) = 0 Then
    MsgBox ("Effective Date is required")
    dlpEDate.SetFocus
    Exit Function
End If

If Val(medPPct.Text) > 1 Then
    MsgBox ("Percentage cannot exceed 100%")
    medPPct.SetFocus
    Exit Function
End If

If clpOHRSDept.Caption = "Unassigned" Then
    MsgBox "OHRS Department # must be valid"
    clpOHRSDept.SetFocus
    Exit Function
End If


chkGLDist = True

Exit Function

chkGLDist_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkGLDist", "HRGLDIST", "edit/Add")
Call RollBack '28July99 js

End Function

Sub cmdCancel_Click()
Dim X
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cancel Error", "HRGLDIST", "Cancel")
Call RollBack '28July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()

Unload Me

If glbOnTop = "frmEGLDist" Then glbOnTop = ""

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
If Not glbtermopen Then
    If Not AUDITBENF("D", 1) Then MsgBox "ERROR - AUDIT FILE"
End If

If glbtermopen Then
    Call Employee_GL_Dist_Integration(glbTERM_ID, , True, glbTERM_Seq, clpGLNbr.Text) 'Ticket #24518 Franks 04/21/2014

    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    Call Employee_GL_Dist_Integration(glbLEE_ID, , True, , clpGLNbr.Text) 'Ticket #24518 Franks 04/21/2014

    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If
 
fglbNew = False

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRGLDIST", "Delete")
Call RollBack '28July99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call SET_UP_MODE

Actn = "M"
oGL = clpGLNbr
oPCT = medPPct
oDATE = dlpEDate
oOHRSDept = clpOHRSDept

fglbNew = False

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRGLDIST", "Modify")
Call RollBack '28July99 js

End Sub

Sub cmdNew_Click()
Dim SQLQ As String
Dim rs As New ADODB.Recordset

fglbNew = True

'Call ST_UPD_MODE(True)
Call SET_UP_MODE


clpGLNbr.SetFocus

On Error GoTo AddN_Err


Call Set_Control("B", Me)

rsDATA.AddNew

If glbtermopen Then
    lblEEID = glbTERM_ID
    SQLQ = "SELECT ED_GLNO FROM Term_HREMP WHERE TERM_SEQ=" & glbTERM_ID
    rs.ActiveConnection = gdbAdoIhr001X
Else
    lblEEID = glbLEE_ID
    SQLQ = "SELECT ED_GLNO FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
    rs.ActiveConnection = gdbAdoIhr001
End If
If Data1.Recordset.RecordCount < 1 And Len(clpGLNbr.Text) = 0 Then
    rs.Open SQLQ, , adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False Then
        If Len(rs("ED_GLNO")) > 0 Then
            clpGLNbr.Text = rs("ED_GLNO")
            medPPct.Text = 1
        End If
    End If
    rs.Close
End If
    
lblCNum.Caption = "001"

Actn = "A"
oGL = ""
oPCT = ""
oDATE = ""
fglbNew = True

exH:
Set rs = Nothing

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRGLDIST", "Add")
Call RollBack '28July99 js
Resume exH
End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X, xID
Dim rsCOM As New ADODB.Recordset

On Error GoTo Add_Err

If Not chkGLDist() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)

If Not glbtermopen Then
    If Not AUDITBENF(Actn, 1) Then MsgBox "ERROR - AUDIT FILE"
End If

Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    rsDATA.Resync
    xID = rsDATA!GL_ID
    Call Employee_GL_Dist_Integration(glbTERM_ID, , , glbTERM_Seq, clpGLNbr.Text) 'Ticket #24518 Franks 04/21/2014
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    rsDATA.Resync
    xID = rsDATA!GL_ID
    Call Employee_GL_Dist_Integration(glbLEE_ID, , , , clpGLNbr.Text) 'Ticket #24518 Franks 04/21/2014
End If
Data1.Refresh

Data1.Recordset.Find "GL_ID=" & xID

fglbNew = False

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

If NextFormIF(lStr("G/L Distribution")) Then
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRGLDIST", "Update")
Call RollBack '28July99 js

End Sub

Sub cmdPrint_Click()
Dim RHeading As String, dscGroup$

    RHeading = lblEEName & "'s GL Distribution"
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    Me.vbxCrystal.Destination = 1
    Me.vbxCrystal.Action = 1
    
End Sub

Sub cmdView_Click()
Dim RHeading As String, dscGroup$

    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

    RHeading = lblEEName & "'s GL Distribution"
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    
    Me.vbxCrystal.Destination = 0
    Me.vbxCrystal.Action = 1

    
End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS

'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_GLDist_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("G/L") & " Distribution information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If

If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_HRGLDIST"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * from HRGLDIST"
    SQLQ = SQLQ & " where GL_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY GL_ID " 'GL_PERCENT DESC, GL_GLNO "
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HRGLDIST", "SELECT")
Call RollBack '28July99 js

Exit Function

End Function

Private Sub cmdCalcPer_Click()
    Dim temp As Double
    Dim C As Long
    Dim fs As ADODB.Recordset
    
    If Data1.Recordset.RecordCount > 0 Then
        Set fs = Data1.Recordset.Clone
        temp = 0
        Do Until fs.EOF
            temp = temp + fs("GL_PERCENT")
            fs.MoveNext
        Loop
            
        If temp = 1 Then
            MsgBox lStr("G/L") & " Distribution = 100%"
        Else
            MsgBox "Warning - " & lStr("G/L") & " Distribution = " & CStr(temp * 100) & "%, please adjust distribution"
        End If
    End If
    
    
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
'Me.cmdModify_Click
    glbOnTop = "frmEGLDist"
    
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "frmEGLDist"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "frmEGLDist"
Actn = " "

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
    If glbUserEmpNo = glbLEE_ID And Not gSec_GLDist_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("G/L") & " Distribution information.", vbCritical, "info:HR - Security"
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

Call setCaption(lblGL)

If Len(glbLEE_SName) < 1 Then Exit Sub

Screen.MousePointer = HOURGLASS

Me.vbxTrueGrid.SetFocus

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = lStr("G/L Distribution - ") & Left$(glbLEE_SName, 5)
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

If gSec_Upd_GLDist Then
    Call ST_UPD_MODE(True)
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
    Set frmEGLDist = Nothing 'carmen may 00
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

clpGLNbr.Enabled = TF
medPPct.Enabled = TF
memComments.Enabled = TF
clpOHRSDept.Enabled = TF

fUPMode = TF    ' update mode

End Sub

Private Sub medPPct_GotFocus()

Call SetPanHelp(Me.ActiveControl)

If Len(medPPct) > 0 Then
    medPPct = medPPct * 100
End If

End Sub

Private Sub medPPct_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = Asc(vbBack) Or IsNumeric(Chr(KeyAscii))) And KeyAscii <> 46 Then KeyAscii = 0
End Sub

Private Sub medPPct_LostFocus()
If (Not IsNumeric(medPPct)) And medPPct.DataChanged Then medPPct = 0
If Len(medPPct) > 0 Then
    medPPct = medPPct / 100
End If
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
            SQLQ = "Select * from Term_HRGLDIST"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HRGLDIST"
            SQLQ = SQLQ & " where GL_EMPNBR = " & glbLEE_ID
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRGLDIST", "Add")
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
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        Call SET_UP_MODE
        
        Me.cmdModify_Click
        
        Exit Sub
    End If
    
    
If glbtermopen Then
    SQLQ = "Select * from Term_HRGLDIST"
    SQLQ = SQLQ & " WHERE GL_ID = " & Data1.Recordset!GL_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HRGLDIST"
    SQLQ = SQLQ & " where GL_ID = " & Data1.Recordset!GL_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
    
Call SET_UP_MODE

clpGLNbr.TextBoxWidth = 1500

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
    UpdateRight = gSec_Upd_GLDist
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
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    Me.Caption = lStr("G/L Distribution - ") & Left$(glbLEE_SName, 5)
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

Private Function AUDITBENF(ACTX, aType)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDITBENF = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, "
strFields = strFields & "AU_DEPT_GL, AU_GL_EDATE, AU_GL_PERCENT, "
strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If ACTX = "D" Then

    rsTA("AU_DEPT_GL") = clpGLNbr
    If IsNumeric(medPPct) Then
    rsTA("AU_GL_PERCENT") = medPPct
    End If
    If IsDate(dlpEDate.Text) Then
        rsTA("AU_GL_EDATE") = dlpEDate.Text
    Else                              '
        rsTA("AU_GL_EDATE") = Null        '
    End If                            '

Else
    If oGL <> clpGLNbr Then GoTo NLine1
    If oPCT <> medPPct Then GoTo NLine1
    If oDATE <> dlpEDate Then GoTo NLine1
    GoTo NLine2
NLine1:
    rsTA("AU_DEPT_GL") = clpGLNbr
    If IsNumeric(medPPct) Then
    rsTA("AU_GL_PERCENT") = medPPct
    End If
    rsTA("AU_GL_EDATE") = dlpEDate.Text
NLine2:
End If

Dim rsEMP As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_DOH FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEMP.EOF Then
    If Not IsNull(rsEMP("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEMP("ED_PAYROLL_ID")
End If
'rsEMP.Close

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = glbLEE_ID
'If Actn = "A" Then
'    rsTA("AU_LDATE") = Format(dlpEDate.Text, "SHORT DATE")
'Else
'    rsTA("AU_LDATE") = Date
'End If
rsTA("AU_LDATE") = Date
'Ticket #23407 Franks 03/15/2013 - begin
If IsDate(rsEMP("ED_DOH")) Then
    If CVDate(rsEMP("ED_DOH")) > CVDate(Date) Then
        rsTA("AU_LDATE") = rsEMP("ED_DOH")
    End If
End If
rsEMP.Close
'Ticket #23407 Franks 03/15/2013 - end
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

MODNOUPD:
AUDITBENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

