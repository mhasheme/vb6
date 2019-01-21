VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmBasicSyndesis 
   Caption         =   "Demographics Information"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6315
      _Version        =   65536
      _ExtentX        =   11139
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Height          =   195
         Left            =   300
         TabIndex        =   23
         Top             =   180
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
         Left            =   1470
         TabIndex        =   22
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         DataField       =   "ED_EMPNBR"
         DataSource      =   "Data1"
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   5250
         TabIndex        =   21
         Top             =   150
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   3060
         TabIndex        =   20
         Top             =   150
         Width           =   1740
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   24
      Top             =   3375
      Width           =   6315
      _Version        =   65536
      _ExtentX        =   11139
      _ExtentY        =   979
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
         Left            =   360
         TabIndex        =   7
         Tag             =   "Close and exit this screen"
         Top             =   30
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
         Left            =   1200
         TabIndex        =   8
         Tag             =   "Edit information on this screen"
         Top             =   30
         Width           =   735
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
         Left            =   2040
         TabIndex        =   9
         Tag             =   "Save changes made"
         Top             =   30
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
         Left            =   2880
         TabIndex        =   10
         Tag             =   "Cancel changes made"
         Top             =   30
         Width           =   795
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8490
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   3
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   4080
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ConnectMode     =   3
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
         Caption         =   "Ado1"
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
   Begin VB.Frame frmBasic 
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   6165
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_LUSER"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   5700
         MaxLength       =   25
         TabIndex        =   27
         Top             =   1620
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_LTIME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   5340
         MaxLength       =   25
         TabIndex        =   26
         Top             =   1620
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox Updstats 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataField       =   "ED_LDATE"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   4980
         MaxLength       =   25
         TabIndex        =   25
         Top             =   1620
         Visible         =   0   'False
         Width           =   330
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   2
         Left            =   1860
         TabIndex        =   0
         Tag             =   "00-Enter Region Code"
         Top             =   270
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDRG"
      End
      Begin INFOHR_Controls.CodeLookup clpGLNum 
         DataField       =   "ED_GLNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1860
         TabIndex        =   1
         Tag             =   "00-General Ledger - Code"
         Top             =   630
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   3
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         DataField       =   "ED_DIV"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1860
         TabIndex        =   2
         Tag             =   "00-Specific Division Desired"
         Top             =   960
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_LOC"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   1
         Left            =   1860
         TabIndex        =   3
         Tag             =   "00-Enter Location Code"
         Top             =   1290
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "ED_ADMINBY"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   3
         Left            =   1860
         TabIndex        =   4
         Tag             =   "00-Enter Administered By Code"
         Top             =   1620
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDAB"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         DataField       =   "ED_DEPTNO"
         DataSource      =   "Data1"
         Height          =   285
         Left            =   1860
         TabIndex        =   5
         Tag             =   "00-Specific Department Desired"
         Top             =   1980
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   4
         Left            =   1860
         TabIndex        =   6
         Tag             =   "00-Enter Section Code"
         Top             =   2340
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDSE"
      End
      Begin VB.Label lblCodeDesc 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   4470
         TabIndex        =   28
         Top             =   285
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Section"
         Height          =   195
         Index           =   26
         Left            =   450
         TabIndex        =   18
         Top             =   2370
         Width           =   540
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   450
         TabIndex        =   17
         Top             =   1050
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   23
         Left            =   450
         TabIndex        =   16
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Region"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   24
         Left            =   450
         TabIndex        =   15
         Top             =   360
         Width           =   510
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Administered By"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   25
         Left            =   450
         TabIndex        =   14
         Top             =   1710
         Width           =   1125
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
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
         Height          =   195
         Index           =   11
         Left            =   450
         TabIndex        =   13
         Top             =   2040
         Width           =   990
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "G/L #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   450
         TabIndex        =   12
         Top             =   720
         Width           =   435
      End
   End
End
Attribute VB_Name = "frmBasicSyndesis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ODept As String, ODeptD As String
Dim OGLNum As String, OGLNumD As String
Dim SavDiv, SavDept, oldEEId
Dim OSection
Dim oRegion, oAdminBy
Dim SavLoc  'laura nov 4, 1997
Dim UnloadForm As Boolean 'Jaddy 10/29/99
Dim RDept, RGLNum ''added by Jaddy Sep 20,99
Dim flagFrmLoad As Boolean   'carmen may 00

Private Sub clpDept_Change()
    If Not cmdOK.Enabled Then RDept = clpDept    'added by Jaddy Sep 20,99
    Call Dept_GL
End Sub

Private Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

RDept = SavDept

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Data1.Refresh
Data1.Refresh
If Not Data1.Recordset.EOF Then Call getCodes


Call ST_UPD_MODE(False)

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
Call RollBack '21June99 js

End Sub

Private Sub cmdClose_Click()
Call NextForm
Unload Me
End Sub

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call ST_UPD_MODE(True)

oGLNo = clpGLNum.Text
oRegion = clpCode(2).Text
oAdminBy = clpCode(3).Text
OSection = clpCode(4).Text

SavDept = clpDept.Text
SavDiv = clpDiv.Text
SavLoc = clpCode(1).Text

clpCode(2).Enabled = True
clpCode(2).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '21June99 js

End Sub

Private Sub cmdOK_Click()
Dim rc%, DtTm As Variant, x%
Dim xDept, xDiv, ctylist
Dim xDeptEDate, xDivEDate
On Error GoTo Add_Err
DtTm = Now
If Not chk_FEBASIC() Then Exit Sub
Screen.MousePointer = HOURGLASS

If SavDept <> clpDept.Text Then xDept = clpDept.Text Else xDept = ""
If SavDiv <> clpDiv.Text Then xDiv = clpDiv.Text Else xDiv = "*"
Call UpdUStats(Me)
If Not glbtermopen Then
    If SavDept <> clpDept.Text Then
        If Not EmpHisCalc(2, glbLEE_ID, xDept, "", "", "", "", "", "", Date) Then MsgBox "EMPHIS Error "
    End If
    If SavDiv <> clpDiv.Text Then
        If Not EmpHisCalc(2, glbLEE_ID, "", xDiv, "", "", "", "", "", Date) Then MsgBox "EMPHIS Error "
    End If
    If SavLoc <> clpCode(1) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "LOC", clpCode(1)) Then MsgBox "EMPHIS Error "
    If oRegion <> clpCode(2) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "REGION", clpCode(2)) Then MsgBox "EMPHIS Error "
    If oAdminBy <> clpCode(3) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "ADMINBY", clpCode(3)) Then MsgBox "EMPHIS Error "
    If OSection <> clpCode(4) Then If Not EmpHisCalc(2, glbLEE_ID, "", "", "", "", "", "", "", Date, "SECTION", clpCode(4)) Then MsgBox "EMPHIS Error "

    If Not AUDITDEMO("M") Then MsgBox "ERROR : AUDIT FILE"
End If
Call UpdCodes
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call ST_UPD_MODE(False)
Screen.MousePointer = DEFAULT
Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMP", "Update")
Call RollBack '21June99 js

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg$, Title  ' Declare variables.
Dim RFound As Integer, VReturn%, x%, xPIC
glbOnTop = "FRMBASICSYNDESIS"
flagFrmLoad = True

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data1.RecordSource = "SELECT " & FldList & " FROM Term_HREMP"
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT " & FldList & " FROM HREMP"
End If
'Data2.ConnectionString = glbAdoIHRDB

Screen.MousePointer = HOURGLASS
Call setCaption(lblTitle(2))
Call setCaption(lblTitle(11))
Call setCaption(lblTitle(12))
Call setCaption(lblTitle(13))
Call setCaption(lblTitle(15))
Call setCaption(lblTitle(23))
Call setCaption(lblTitle(24))
Call setCaption(lblTitle(25))
Call setCaption(lblTitle(26))

Screen.MousePointer = DEFAULT




UnloadForm = False
If Not glbtermopen Then
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

If Not EERetrieve() Then
    MsgBox "Sorry, Employee can not be found"
    Exit Sub
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

Call getCodes
Call ST_UPD_MODE(False)
If Not gSec_Upd_Basic Then         'May99 js
    Call ST_UPD_MODE(False)      '
    cmdModify.Enabled = False   '
    cmdNew.Enabled = False       '
    cmdDelete.Enabled = False
End If                          '
Call setCaption(lblTitle(15))
MDIMain.panHelp(0).Caption = "Demographics Information"
Call INI_Controls(Me)
End Sub

Sub getCodes()
    If Not IsNull(Data1.Recordset("ED_REGION")) Then
        clpCode(2).Text = Data1.Recordset("ED_REGION")
    Else
        clpCode(2).Text = ""
    End If
    If Not IsNull(Data1.Recordset("ED_SECTION")) Then
        clpCode(4).Text = Data1.Recordset("ED_SECTION")
    Else
        clpCode(4).Text = ""
    End If
End Sub


Private Sub CR_Lgr_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo Job_Err

Screen.MousePointer = HOURGLASS
SQLQ = "SELECT * FROM HRGL;"

If LGR_snap.State <> 0 Then LGR_snap.Close
LGR_snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

Screen.MousePointer = DEFAULT

Exit Sub

Job_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Descriptions", "HRGL", "SELECT")
Call RollBack '21June99 js

End Sub
Private Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
If glbtermopen Then
    SQLQ = "Select " & FldList & " from Term_HREMP"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select " & FldList & " from HREMP"
    SQLQ = SQLQ & " where ED_EMPNBR = " & glbLEE_ID
End If
Data1.RecordSource = SQLQ
Data1.Refresh


EERetrieve = True

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack '21June99 js

Exit Function

End Function

Private Function FldList()
Dim SQLQ
SQLQ = ""
SQLQ = SQLQ & "ED_EMPNBR,ED_SURNAME,ED_FNAME,ED_DEPTNO,ED_GLNO,ED_PT,"
SQLQ = SQLQ & "ED_DIV, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION,"
SQLQ = SQLQ & "ED_HOMELINE,ED_HOMESHIFT,ED_HOMEOPRTNBR,ED_HOMEWRKCNT,"
SQLQ = SQLQ & "ED_LDATE, ED_LTIME,ED_LUSER"
If glbtermopen Then SQLQ = SQLQ & ",TERM_SEQ"
FldList = SQLQ
End Function

Private Function AUDITDEMO(Actn)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xProvNbr, xADD

On Error GoTo AUDIT_ERR

AUDITDEMO = False



rsTB.Open "HRPROV", gdbAdoIhr001, adOpenKeyset, , adCmdTableDirect
rsTB.Find "CODE = '" & txtProv & "'"

If rsTB.EOF Then
    xProvNbr = "  "
Else
    xProvNbr = rsTB("NBR")
End If


rsTA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect

xADD = False

If Actn = "M" Then
    If SavDiv <> clpDiv.Text Or SavDept <> clpDept.Text Then GoTo MODUPD
    If oGLNo <> clpGLNum.Text Or OMSTAT <> txtMStatus Then GoTo MODUPD
    If SavLoc <> clpCode(1).Text Then GoTo MODUPD   'Laura nov 4, 1997
    If oRegion <> clpCode(2).Text Then GoTo MODUPD
    If oAdminBy <> clpCode(3).Text Then GoTo MODUPD
    If OSection <> clpCode(4).Text Then GoTo MODUPD
    GoTo MODNOUPD
MODUPD:
    xADD = True
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN": rsTA("AU_UPLOAD") = "N"
    rsTA("AU_DIVUPL") = Data1.Recordset("ED_DIV")
    rsTA("AU_TYPE") = "M"
    rsTA("AU_PTUPL") = Data1.Recordset("ED_PT") ' added by jrowland 9/19/97
    rsTA("AU_NEWEMP") = "N"
    If SavDiv <> clpDiv.Text Then
        rsTA("AU_OLDDIV") = SavDiv
        rsTA("AU_DIV") = clpDiv.Text
        rsTA("AU_DIVUPL") = clpDiv.Text
    End If
    If SavDept <> clpDept.Text Then
        rsTA("AU_OLDDEPT") = SavDept
        rsTA("AU_DEPTNO") = clpDept.Text
    End If
    If SavLoc <> clpCode(1).Text Then   'laura nov 4, 1997
        If SavLoc <> "" Then rsTA("AU_OLDLOC") = SavLoc
        If clpCode(1).Text <> "" Then rsTA("AU_LOC") = clpCode(1).Text
    End If
    If oRegion <> clpCode(2).Text Then rsTA("AU_REGION") = clpCode(2).Text
    If oAdminBy <> clpCode(3).Text Then rsTA("AU_ADMINBY") = clpCode(3).Text
    If OSection <> clpCode(4).Text Then rsTA("AU_SECTION") = clpCode(4).Text
    If ODeptEDate <> txtDeptEDate Then If IsDate(txtDeptEDate) Then rsTA("AU_DEPTEDATE") = txtDeptEDate
    If ODivEdate <> txtDivEDate Then If IsDate(txtDivEDate) Then rsTA("AU_DIVEDATE") = txtDivEDate
MODNOUPD:
End If

If xADD Then
    rsTA("AU_PTUPL") = "FT" 'added by jrowland 9/19/97
    rsTA("AU_DIVUPL") = clpDiv.Text
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = Actn
    rsTA.Update
End If

AUDITDEMO = True

Exit Function
AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '18June99 js

End Function
Private Sub ST_UPD_MODE(YN)
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

cmdClose.Enabled = FT
cmdModify.Enabled = FT

clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
 clpDept.Enabled = TF
 clpDiv.Enabled = TF And Not glbLinamar
 clpGLNum.Enabled = TF
If glbtermopen Then
    cmdNew.Enabled = False
End If
End Sub

Private Sub UpdCodes()
        Data1.Recordset("ED_REGION") = clpCode(2).Text
        Data1.Recordset("ED_SECTION") = clpCode(4).Text
End Sub

Private Function chk_FEBASIC()
Dim VReturn%
Dim mSIN As String
Dim EditFlag
Dim x

EditFlag = True

If glbCompSerial = "S/N - 2207W" And txtProv <> "ON" Then 'ADDED BY RAUBREY 6/2/97
  EditFlag = False
End If

chk_FEBASIC = False
If Len(clpDept.Text) < 1 Then
    MsgBox lStr("Department is a required field")
     clpDept.SetFocus
    Exit Function
Else
    If clpDept.Caption = "Unassigned" Then
        MsgBox "Department Code must be valid"
        Call Get_Dept(False)
        If Len(glbDept) < 1 Then
             clpDept.Text = ODept
             clpGLNum.Text = OGLNum
             clpDept.Caption = ODeptD
             clpGLNum.Caption = OGLNumD
             clpDept.ShowDescription = True
             clpGLNum.ShowDescription = True
        Else
             clpDept.Text = glbDept
             clpGLNum.Text = glbGLNum
             clpDept.Caption = glbDeptDesc
        End If
        fglb_FindDept = False
        Exit Function
    End If
End If

If Len(clpDiv.Text) > 1 And clpDiv.Caption = "Unassigned" Then
    If Not glbLinamar Then
        MsgBox lStr("If Division is entered it must be valid")
         clpDiv.SetFocus
        Exit Function
    End If
End If

If Len(clpGLNum.Text) > 0 And clpGLNum.Caption = "Unassigned" Then
    MsgBox "If G/L Number is entered it must be valid"
     clpGLNum.SetFocus
    Exit Function
End If
For x = 1 To 4
    If Len(clpCode(x).Text) > 0 And clpCode(x).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(x).SetFocus
        Exit Function
    End If
Next x

chk_FEBASIC = True

End Function


Private Sub lblEEID_Change()

Caption = "Demographics Information - " & Data1.Recordset("ED_SURNAME")
frmBasicLinamar.lblEEName = RTrim$(Data1.Recordset("ED_SURNAME")) & ", " & RTrim$(Data1.Recordset("Ed_FNAME"))
lblEENum = ShowEmpnbr(lblEEID)
End Sub

Private Sub Dept_GL()
Dim Response%, Msg$, Title$, DgDef As Double
Dim SQLQ As String
Dim rsDEPT As New ADODB.Recordset
On Error GoTo Dept_GL_Err

If Len(clpDept.Text) > 0 Then
    rsDEPT.Open "SELECT DF_GLNO FROM HRDEPT WHERE DF_NBR='" & clpDept.Text & "'", gdbAdoIhr001
    If Not rsDEPT.EOF Then
        RGLNum = rsDEPT("DF_GLNO")
        If RDept <> clpDept Then
            If IsNull(RGLNum) Then
                RGLNum = ""
            Else
                Msg$ = "Do you want the associated G/L #?"
                Title$ = "info:HR"
                DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
                Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
                If Response% = IDYES Then clpGLNum.Text = RGLNum
            End If
            RDept = clpDept.Text
        End If
    End If
End If

Exit Sub

Dept_GL_Err:
If Err = 94 Then
     clpGLNum.Text = ""
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Dept Snap", "DEPT", "SELECT")
Call RollBack '21June99 js
End Sub




Private Sub Div_Desc()
Dim SQLQ As String

On Error GoTo clpDiv_Err
If Len(clpDiv.Text) > 0 Then
    SQLQ = "DIV = '" & clpDiv.Text & "'"
    snapDiv.Requery
    snapDiv.Find SQLQ
    If snapDiv.EOF Then
         clpDiv.Caption = "Unassigned"
    Else
         clpDiv.Caption = snapDiv("Division_Name")
    End If
     clpDiv.ShowDescription = True
Else
     clpDiv.Caption = ""
     clpDiv.ShowDescription = False
End If

Exit Sub

clpDiv_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Div Snap", "Div", "SELECT")
Call RollBack '21June99 js

End Sub

Private Sub LGR_Desc()

Dim SQLQ As String

On Error GoTo Jobd_Err
If Len(clpGLNum.Text) > 0 Then
    SQLQ = "GL_NO = '" & CStr(clpGLNum.Text) & "'"
    LGR_snap.Requery
    LGR_snap.Find SQLQ
    
    If LGR_snap.EOF Then
         clpGLNum.Caption = "Unassigned"
         clpGLNum.ShowDescription = False
        Exit Sub
    End If
     clpGLNum.Caption = LGR_snap("GL_DESCR")
     clpGLNum.ShowDescription = True
Else
     clpGLNum.ShowDescription = False
End If

Exit Sub

Jobd_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "LEAGERS", "SELECT")
Call RollBack '21June99 js

End Sub





