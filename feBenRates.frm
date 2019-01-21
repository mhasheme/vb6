VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSBenRates 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Benefit Rates"
   ClientHeight    =   6480
   ClientLeft      =   90
   ClientTop       =   1005
   ClientWidth     =   9495
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
   ScaleHeight     =   6480
   ScaleWidth      =   9495
   WindowState     =   2  'Maximized
   Begin VB.Frame frmSex 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1710
      TabIndex        =   16
      Top             =   3765
      Width           =   1995
      Begin VB.OptionButton optGender 
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1050
         TabIndex        =   4
         Tag             =   "41-Gender"
         Top             =   30
         Width           =   930
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Tag             =   "41-Gender"
         Top             =   30
         Width           =   675
      End
   End
   Begin VB.TextBox txtSmoker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "BR_SMOKER"
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
      Left            =   2910
      MaxLength       =   2
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Text14"
      Top             =   4170
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.ComboBox comSmoker 
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
      Height          =   315
      Left            =   1935
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "01-Smoker Yes/No"
      Top             =   4155
      Width           =   855
   End
   Begin VB.TextBox txtGender 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "BR_SEX"
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
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "Text14"
      Top             =   3765
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox txtBenType 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      DataField       =   "BR_BENTYPE"
      Height          =   285
      Left            =   3300
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4995
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.ComboBox cmbBenType 
      Height          =   315
      ItemData        =   "feBenRates.frx":0000
      Left            =   1920
      List            =   "feBenRates.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Tag             =   "41-Select Benefit Type"
      Top             =   4980
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4920
      Top             =   5880
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DG_LUSER"
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
      Left            =   8280
      MaxLength       =   10
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DG_LDATE"
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
      Index           =   0
      Left            =   6840
      MaxLength       =   12
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3465
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "DG_LTIME"
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
      Left            =   7590
      MaxLength       =   8
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3480
      Visible         =   0   'False
      Width           =   645
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   4320
      Top             =   5880
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
      GridSource      =   "vbxTrueGrid"
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feBenRates.frx":0028
      Height          =   2775
      Left            =   240
      OleObjectBlob   =   "feBenRates.frx":003C
      TabIndex        =   0
      Top             =   240
      Width           =   8895
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "BR_BCODE"
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   8
      Tag             =   "01-Benefit - Code"
      Top             =   5400
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BNCD"
      MaxLength       =   10
   End
   Begin MSMask.MaskEdBox medRate 
      DataField       =   "BR_RATE"
      Height          =   285
      Left            =   1935
      TabIndex        =   6
      Tag             =   "21-Enter Rate"
      Top             =   4590
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "$#,##0.000000;($#,##0.000000)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medAgeFrom 
      DataField       =   "BR_AGE_FROM"
      Height          =   285
      Left            =   1935
      TabIndex        =   1
      Tag             =   "11-From Age"
      Top             =   3360
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
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
   Begin MSMask.MaskEdBox medAgeTo 
      DataField       =   "BR_AGE_TO"
      Height          =   285
      Left            =   3000
      TabIndex        =   2
      Tag             =   "11-To Age"
      Top             =   3360
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      PromptInclude   =   0   'False
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
   Begin VB.Label lblAgeFrom 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Age From"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   23
      Top             =   3405
      Width           =   810
   End
   Begin VB.Label lblAgeTo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2640
      TabIndex        =   22
      Top             =   3405
      Width           =   240
   End
   Begin VB.Label lblGender 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   3810
      Width           =   630
   End
   Begin VB.Label lblSmoker 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Smoker?"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   20
      Top             =   4215
      Width           =   750
   End
   Begin VB.Label lblRate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   19
      Top             =   4635
      Width           =   420
   End
   Begin VB.Label lblBenefit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   18
      Top             =   5445
      Width           =   615
   End
   Begin VB.Label lblBenType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   360
      TabIndex        =   17
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "DG_COMPNO"
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
      Left            =   7680
      TabIndex        =   11
      Top             =   6000
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmSBenRates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fGLBNew As Boolean
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim DefType(0 To 3)
Dim SystType(0 To 3)
Dim RSDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum

Private Function chkBenRates()
Dim Msg As String
Dim X%
Dim rsBenRates As New ADODB.Recordset
Dim SQLQ As String

chkBenRates = False

If Len(Trim(medAgeFrom)) = 0 Then
    MsgBox "Age From is required", 48
    medAgeFrom.SetFocus
    Exit Function
End If

If Not IsNumeric(medAgeFrom) Then
    MsgBox "Age From is invalid", 48
    medAgeFrom.SetFocus
    Exit Function
End If

If Len(Trim(medAgeTo)) = 0 Then
    MsgBox "Age To is required", 48
    medAgeTo.SetFocus
    Exit Function
End If

If Not IsNumeric(medAgeTo) Then
    MsgBox "Age To is invalid", 48
    medAgeTo.SetFocus
    Exit Function
End If

If Len(Trim(medRate)) = 0 Then
    MsgBox "Rate is required", 48
    medRate.SetFocus
    Exit Function
End If

If Not IsNumeric(medRate) Then
    MsgBox "Rate is invalid", 48
    medRate.SetFocus
    Exit Function
End If

If cmbBenType.ListIndex = -1 Then
    MsgBox "Benefit Type is required", 48
    cmbBenType.SetFocus
    Exit Function
End If

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Benefit code is a required field", 48
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Benefit code must be valid", 48
    clpCode(1).SetFocus
    Exit Function
End If


If optGender(0).Value = True Then
    txtGender = "M"
Else
    txtGender = "F"
End If
txtBenType = Left(cmbBenType, 1)

'Check if the duplicate Benefit Rate already exists
SQLQ = "SELECT * FROM HR_BENEFIT_RATES"
SQLQ = SQLQ & " WHERE BR_AGE_FROM = " & medAgeFrom
SQLQ = SQLQ & " AND BR_AGE_TO = " & medAgeTo
SQLQ = SQLQ & " AND BR_SEX = '" & txtGender.Text & "'"
SQLQ = SQLQ & " AND BR_SMOKER = " & IIf(comSmoker = "Yes", 1, 0)
SQLQ = SQLQ & " AND BR_BENTYPE = '" & txtBenType.Text & "'"
SQLQ = SQLQ & " AND BR_BCODE = '" & clpCode(1).Text & "'"
'SQLQ = SQLQ & " AND BR_RATE = " & medRate.Text
If Not fGLBNew Then
    SQLQ = SQLQ & " AND BR_ID <> " & Data1.Recordset!BR_ID
End If
SQLQ = SQLQ & " ORDER BY BR_AGE_FROM,BR_AGE_TO,BR_SEX,BR_SMOKER,BR_BENTYPE,BR_BCODE"
rsBenRates.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsBenRates.EOF Then
    rsBenRates.Close
    Set rsBenRates = Nothing
Else
    MsgBox "This Benefit Rate already exists."
    rsBenRates.Close
    Set rsBenRates = Nothing
    medAgeFrom.SetFocus
    Exit Function
End If


chkBenRates = True

End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

fGLBNew = False

If fglbEmptyNew Then
    Me.vbxTrueGrid.Enabled = True
    Me.vbxTrueGrid.Refresh
End If

RSDATA.CancelUpdate

Call Display_Value


'Call ST_UPD_MODE(True) ' reset screen's attributes

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_BENEFIT_RATES", "Cancel")
Call RollBack '09June99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim X As Integer
Dim xID As Integer

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
If vbxTrueGrid.SelBookmarks.count > 1 Then
    Msg = Msg & "These Records?"
Else
    Msg = Msg & "This Record?"
End If
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

'7.9 Enhancement to delete multiple records.
If vbxTrueGrid.SelBookmarks.count = 0 Then vbxTrueGrid.SelBookmarks.Add Data1.Recordset.Bookmark
For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
    Data1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(X)
    xID = Data1.Recordset("BR_ID")
    
    gdbAdoIhr001.BeginTrans
    'rsDATA.Delete
    gdbAdoIhr001.Execute "DELETE FROM HR_BENEFIT_RATES WHERE BR_ID=" & xID
    gdbAdoIhr001.CommitTrans
    DoEvents
Next

Data1.Refresh


Call SET_UP_MODE
'Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_BENEFIT_RATES", "Delete")
Call RollBack '09June99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Call ST_UPD_MODE(True)

medAgeFrom.SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_BENEFIT_RATES", "Modify")
Call RollBack '09June99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()

On Error GoTo AddN_Err

Call Set_Control("B", Me)

RSDATA.AddNew

lblCNum.Caption = "001"
txtGender = "M"

fGLBNew = True

Call SET_UP_MODE

'Call ST_UPD_MODE(True)
medAgeFrom.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_BENEFIT_RATES", "Add")
Call RollBack '09June99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X%
Dim bmk As Variant

On Error GoTo cmdOK_Err
If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

If Not chkBenRates() Then Exit Sub

If comSmoker = "Yes" Then txtSmoker = "-1" Else txtSmoker = "0"
txtBenType = Left(cmbBenType, 1)
If optGender(0).Value = True Then
    txtGender = "M"
Else
    txtGender = "F"
End If

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, RSDATA)

gdbAdoIhr001.BeginTrans
RSDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
If Not bmk = 0 Then
    Data1.Recordset.Bookmark = bmk
End If

fGLBNew = False

Call Display_Value

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_BENEFIT_RATES", "Update")
Call RollBack '09June99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = "Benefit Rates"
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

RHeading = "Benefit Rates"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbBenType_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbBenType_LostFocus()
    txtBenType = Left(cmbBenType, 1)
End Sub

Private Sub comSmoker_Click()
If comSmoker = "Yes" Then txtSmoker = "-1" Else txtSmoker = "0"
End Sub

Private Sub comSmoker_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_BENEFIT_RATES", "SELECT")

End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%, SQLQ


Me.Show

glbOnTop = "FRMSBENRATES"

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB

Data1.RecordSource = "SELECT * FROM HR_BENEFIT_RATES ORDER BY BR_AGE_FROM,BR_AGE_TO,BR_SEX,BR_SMOKER,BR_BENTYPE,BR_BCODE"

Data1.Refresh

'Call setCaption(lblDept)
'Call setCaption(lblTitle(12))

comSmoker.Clear
comSmoker.AddItem "No"
comSmoker.AddItem "Yes"
comSmoker.ListIndex = 0

cmbBenType.Clear
cmbBenType.AddItem "Employee"
cmbBenType.AddItem "Spousal"
'cmbBenType.AddItem "Child"
cmbBenType.AddItem "Other"
cmbBenType.ListIndex = -1


Screen.MousePointer = DEFAULT

'Call Display_Value

Call ST_UPD_MODE(False)

If Not gSec_BenefitGroupSetup Then  ' gSec_Upd_BenRates                                  'May99 js
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If


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
Dim I As Integer
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

fUPMode = TF
'vbxTrueGrid.Enabled = FT
'cmdOK.Enabled = TF              '
'cmdCancel.Enabled = TF          '
'cmdClose.Enabled = FT           '
'cmdModify.Enabled = FT          '
'cmdNew.Enabled = FT             '
'cmdDelete.Enabled = FT          '
'cmdPrint.Enabled = FT           '


medAgeFrom.Enabled = TF                    '
medAgeTo.Enabled = TF
frmSex.Enabled = TF
comSmoker.Enabled = TF
medRate.Enabled = TF
cmbBenType.Enabled = TF
clpCode(1).Enabled = TF

If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
 '   cmdDelete.Enabled = False
End If

End Sub

Private Sub medAgeFrom_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medAgeTo_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub medRate_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub optGender_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub optGender_KeyPress(Index As Integer, KeyAscii As Integer)
    If optGender(0).Value = True Then
        txtGender = "M"
    Else
        txtGender = "F"
    End If
End Sub

Private Sub optGender_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optGender(0).Value = True Then
        txtGender = "M"
    Else
        txtGender = "F"
    End If
End Sub

Private Sub txtBenType_Change()
    cmbBenType.ListIndex = -1
    Select Case txtBenType
    Case "E"
        cmbBenType.ListIndex = 0
    Case "S"
        cmbBenType.ListIndex = 1
    'Case "C"
    '    cmbBenType.ListIndex = 1
    Case "O"
        cmbBenType.ListIndex = 2
    End Select
End Sub

Private Sub txtGender_Change()
    If Len(txtGender) > 0 Then
        If txtGender = "M" Then
            optGender(0) = True
            optGender(1) = False
        Else
            optGender(0) = False
            optGender(1) = True
        End If
    End If
End Sub

Private Sub txtSmoker_Change()
    If txtSmoker = "-1" Then comSmoker.ListIndex = 1 Else comSmoker.ListIndex = 0
End Sub

'Private Sub txtDept_KeyPress(KeyAscii As Integer)
'KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
'End Sub

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
        
        SQLQ = "SELECT * FROM HR_BENEFIT_RATES "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim I%

On Error GoTo vbxTrueGrid_Err

Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If


Exit Sub

vbxTrueGrid_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_BENEFIT_RATES", "Select")
Call RollBack '09June99 js

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
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_BENEFIT_RATES WHERE BR_ID= " & Data1.Recordset!BR_ID
    
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    
    Call Set_Control("R", Me, RSDATA)
    Call SET_UP_MODE
    
    
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
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_BenefitGroupSetup    'gSec_Upd_BenRates
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
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub


