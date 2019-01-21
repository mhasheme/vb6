VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmELang 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Languages"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9660
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
   Icon            =   "feLanguage.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7905
   ScaleWidth      =   9660
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLangNum 
      Appearance      =   0  'Flat
      DataField       =   "EL_LANGNO"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2610
      MaxLength       =   8
      TabIndex        =   17
      Tag             =   "11- Document Number"
      Top             =   4560
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkWRTN 
      Caption         =   "Written"
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
      Left            =   8160
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox chkSPK 
      Caption         =   "Spoken"
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
      Left            =   6240
      TabIndex        =   13
      Top             =   4920
      Visible         =   0   'False
      Width           =   1335
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid2 
      Bindings        =   "feLanguage.frx":030A
      Height          =   2145
      Left            =   120
      OleObjectBlob   =   "feLanguage.frx":031E
      TabIndex        =   0
      Top             =   600
      Width           =   9015
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6960
      Top             =   7680
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
      TabIndex        =   10
      Top             =   7245
      Width           =   9660
      _Version        =   65536
      _ExtentX        =   17039
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
         Left            =   6495
         Top             =   120
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EL_LDATE"
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
      Left            =   2880
      MaxLength       =   25
      TabIndex        =   1
      Top             =   6300
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EL_LTIME"
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
      Left            =   4560
      MaxLength       =   25
      TabIndex        =   2
      Top             =   6300
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EL_LUSER"
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
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   3
      Top             =   6300
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9660
      _Version        =   65536
      _ExtentX        =   17039
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
         Left            =   6240
         TabIndex        =   22
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   160
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
         TabIndex        =   6
         Top             =   135
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
         TabIndex        =   5
         Top             =   135
         Width           =   720
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EL_LANG_SPOKEN"
      DataSource      =   " "
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   11
      Tag             =   "00-Language - Code"
      Top             =   3000
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDL1"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EL_LANG_WRITTEN"
      DataSource      =   " "
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   15
      Tag             =   "00-Language - Code"
      Top             =   3480
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDL1"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataSource      =   " "
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   19
      Tag             =   "00-Language - Code"
      Top             =   4920
      Visible         =   0   'False
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDL1"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feLanguage.frx":305F
      Height          =   2145
      Left            =   120
      OleObjectBlob   =   "feLanguage.frx":3073
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.Label lblLang1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Language"
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
      Index           =   3
      Left            =   360
      TabIndex        =   20
      Top             =   4920
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblLang1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Language #"
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
      Index           =   2
      Left            =   360
      TabIndex        =   18
      Top             =   4560
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label lblLang1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Language Written"
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
      Index           =   1
      Left            =   360
      TabIndex        =   16
      Top             =   3510
      Width           =   1275
   End
   Begin VB.Label lblLang1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Language Spoken"
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
      Index           =   0
      Left            =   360
      TabIndex        =   12
      Top             =   3030
      Width           =   1320
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EL_EMPNBR"
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
      Left            =   2160
      TabIndex        =   8
      Top             =   6420
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EL_COMPNO"
      DataSource      =   "Data1"
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
      Left            =   330
      TabIndex        =   9
      Top             =   6420
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmELang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer
Dim fglbNewSalRec%
Dim glbNew
Dim fglbNew
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO

Private Function chkLanguages()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, Msg As String
Dim RsLang As New ADODB.Recordset

chkLanguages = False

On Error GoTo chkLanguage_Err

'??? Which one is required field?
If Not glbtermopen Then
    SQLQ = "SELECT * FROM HR_LANGUAGE WHERE EL_EMPNBR = " & glbLEE_ID
Else
    SQLQ = "SELECT * FROM TERM_HR_LANGUAGE WHERE TERM_SEQ = " & glbTERM_Seq
End If
If Not fglbNew Then SQLQ = SQLQ & " AND EL_LANGNO = " & Val(txtLangNum)

SQLQ = SQLQ & " order by EL_LANGNO desc"

RsLang.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not RsLang.EOF Then
    If Not IsNull(RsLang("EL_LANGNO")) Then
        If fglbNew Then
        '    txtLangNum = RsEHST("DE_DOCNO")
        'Else
            txtLangNum = RsLang("EL_LANGNO") + 1
        End If
    End If
Else
    txtLangNum = 1
End If
RsLang.Close

chkLanguages = True

Exit Function

chkLanguage_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkLanguage", "HR_LANGUAGE", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err
'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove ADO
rsDATA.CancelUpdate
Call Display_Value



'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_LANGUAGE", "Cancel")
Call RollBack '23July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMELang" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, X

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

'If Not chkLastLang() Then
'    MsgBox ("Only the last one can be deleted!")
'    Exit Sub
'End If
Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_LANGUAGE", "Delete")
Call RollBack '23July99 js

End Sub




'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew



If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"

Me.clpCode(0).Enabled = True
Me.clpCode(0).SetFocus
Me.clpCode(1).Enabled = True

glbNew = True
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_LANGUAGE", "Add")
Resume Next
End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X
On Error GoTo Add_Err

If Not chkLanguages() Then Exit Sub


Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
If NextFormIF("Language") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    Data1.Recordset.CancelUpdate
    Data1.Recordset.Resync
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_LANGUAGE", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Languages"
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

RHeading = lblEEName & "'s Languages"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_LANGUAGE", "SELECT")

End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS


If glbtermopen Then
    SQLQ = "Select * from Term_HR_LANGUAGE"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * "
    SQLQ = SQLQ & " from HR_LANGUAGE "
    SQLQ = SQLQ & " where EL_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY EL_LANG_SPOKEN,EL_LANG_WRITTEN DESC"
Data1.RecordSource = SQLQ
Data1.Refresh


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LANGsRetrieve", "HR_LANGUAGE", "SELECT")
Resume Next

Exit Function

End Function

Private Sub Form_Activate()
    glbOnTop = "FRMELang"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMELang"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMELang"
If glbtermopen Then  'Lucy July 4, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = HOURGLASS

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




If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Languages - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)
Call Display_Value
Call ST_UPD_MODE(False)
If Not gSec_Upd_Basic Then
'    cmdNew.Enabled = False
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
'If glbLinamar Then clpCode(1).TextBoxWidth = 2000

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT

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
    Call NextForm
End Sub

Private Sub lblDesc_Click()

End Sub

Private Sub medSKLevel_GotFocus()
    Call SetPanHelp(ActiveControl)
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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
 
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If

clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
'txtLangNum.Enabled = TF
'vbxTrueGrid.Enabled = FT

End Sub


Private Sub medSKLevel_KeyPress(KeyAscii As Integer)
If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
    Exit Sub
End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False

End Sub

Private Sub txtSKComment_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtSKDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtSKDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtSKDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtSKDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbtermopen Then
            SQLQ = "Select * from Term_HR_LANGUAGE"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * "
            SQLQ = SQLQ & " from HR_LANGUAGE "
            SQLQ = SQLQ & " where EL_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid2_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid2_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid2_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid2.Tag = "ASC" Then
            vbxTrueGrid2.Tag = "DESC"
        Else
            vbxTrueGrid2.Tag = "ASC"
        End If
        
        If glbtermopen Then
            SQLQ = "Select * from Term_HR_LANGUAGE"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * "
            SQLQ = SQLQ & " from HR_LANGUAGE "
            SQLQ = SQLQ & " where EL_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid2.Columns(ColIndex).DataField & " " & vbxTrueGrid2.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid2_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

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


Private Sub vbxTrueGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Skll As String, Skllvl As String, SklDte As String
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
'Sam add july 2002 * remove ado
Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_LANGUAGE", "Add")
Resume Next

End Sub


Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Skll As String, Skllvl As String, SklDte As String
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
'Sam add july 2002 * remove ado
Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_LANGUAGE", "Add")
Resume Next

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
Else
    If glbtermopen Then
        SQLQ = "Select * from Term_HR_LANGUAGE"
        SQLQ = SQLQ & " WHERE EL_ID = " & Data1.Recordset!EL_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select * "
        SQLQ = SQLQ & " from HR_LANGUAGE "
        SQLQ = SQLQ & " where EL_ID = " & Data1.Recordset!EL_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End If
Call SET_UP_MODE
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
UpdateRight = gSec_Upd_EMP_LANG
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
    frmELang.Caption = "Languages - " & Left$(glbLEE_SName, 5)
    frmELang.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Function chkLastLang()
    Dim rsLastOne As New ADODB.Recordset
    Dim SQLQ
    chkLastLang = False
    If glbtermopen Then
        SQLQ = "Select EL_LANGNO from Term_HR_LANGUAGE"
        SQLQ = SQLQ & " WHERE EL_LANGNO > " & Data1.Recordset!EL_LANGNO
        rsLastOne.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly, adLockReadOnly
    Else
        SQLQ = "Select EL_LANGNO "
        SQLQ = SQLQ & " from HR_LANGUAGE "
        SQLQ = SQLQ & " WHERE EL_LANGNO > " & Data1.Recordset!EL_LANGNO
        rsLastOne.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly, adLockReadOnly
    End If
    If rsLastOne.EOF Or rsLastOne.BOF Then
        chkLastLang = True
    End If
    rsLastOne.Close
    Set rsLastOne = Nothing
End Function



