VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPROV 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Provinces and States"
   ClientHeight    =   7650
   ClientLeft      =   1485
   ClientTop       =   885
   ClientWidth     =   12795
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7650
   ScaleWidth      =   12795
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProvText3 
      Appearance      =   0  'Flat
      DataField       =   "PR_TXT3"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   9
      Tag             =   "01-User Text 3"
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox txtProvText2 
      Appearance      =   0  'Flat
      DataField       =   "PR_TXT2"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   8
      Tag             =   "01-User Text 2"
      Top             =   5520
      Width           =   735
   End
   Begin VB.TextBox txtProvText1 
      Appearance      =   0  'Flat
      DataField       =   "PR_TXT1"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   4
      TabIndex        =   7
      Tag             =   "01-User Text 1"
      Top             =   5160
      Width           =   735
   End
   Begin VB.TextBox txtProvNum2 
      Appearance      =   0  'Flat
      DataField       =   "PR_NUM2"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   6
      Tag             =   "10-User Number 2"
      Top             =   4800
      Width           =   975
   End
   Begin VB.TextBox txtProvNum1 
      Appearance      =   0  'Flat
      DataField       =   "PR_NUM1"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   5
      Tag             =   "10-User Number 1"
      Top             =   4440
      Width           =   975
   End
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "CODE"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   1
      Tag             =   "01-Province Code"
      Top             =   2970
      Width           =   540
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      DataField       =   "DESCR"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   30
      TabIndex        =   2
      Tag             =   "01-Description (Province)"
      Top             =   3360
      Width           =   3975
   End
   Begin VB.TextBox txtComp 
      Appearance      =   0  'Flat
      DataField       =   "COMPNO"
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
      Left            =   12120
      MaxLength       =   3
      TabIndex        =   22
      Top             =   3000
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtNumber 
      Appearance      =   0  'Flat
      DataField       =   "NBR"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   3
      Tag             =   "01-Province #"
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox txtCountry 
      Appearance      =   0  'Flat
      DataField       =   "COUNTRY"
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
      Height          =   285
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "01-Country"
      Top             =   4080
      Width           =   975
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   10800
      Top             =   6720
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
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
      TabIndex        =   21
      Top             =   6990
      Width           =   12795
      _Version        =   65536
      _ExtentX        =   22569
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
         Left            =   60
         TabIndex        =   13
         Tag             =   "Select Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Tag             =   "Close and exit screen"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1635
         TabIndex        =   15
         Tag             =   "Edit the information above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2445
         TabIndex        =   16
         Tag             =   "Save the changes made"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3240
         TabIndex        =   17
         Tag             =   "Cancel the changes made"
         Top             =   165
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4125
         TabIndex        =   18
         Tag             =   "Add a new Province to the list"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4935
         TabIndex        =   19
         Tag             =   "Delete the Province listed above"
         Top             =   165
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5745
         TabIndex        =   20
         Tag             =   "Print the Province listing report"
         Top             =   165
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   5000
      TabIndex        =   12
      Tag             =   "Find specific record"
      Top             =   6570
      Width           =   735
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
      Left            =   840
      TabIndex        =   11
      Tag             =   "00-Search Description"
      Top             =   6600
      Width           =   3975
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
      TabIndex        =   10
      Tag             =   "00-Search Code"
      Top             =   6600
      Width           =   540
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxprov.frx":0000
      Height          =   2745
      Left            =   30
      OleObjectBlob   =   "fxprov.frx":0014
      TabIndex        =   0
      Tag             =   "Province Listings"
      Top             =   0
      Width           =   12615
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   10320
      Top             =   6720
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
   Begin VB.Label lblProvTxt3 
      AutoSize        =   -1  'True
      Caption         =   "Prov. Text3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   5880
      Width           =   825
   End
   Begin VB.Label lblProvTxt2 
      AutoSize        =   -1  'True
      Caption         =   "Prov. Text2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   5520
      Width           =   825
   End
   Begin VB.Label lblProvTxt1 
      AutoSize        =   -1  'True
      Caption         =   "Prov. Text1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   29
      Top             =   5160
      Width           =   825
   End
   Begin VB.Label lblProvNum2 
      AutoSize        =   -1  'True
      Caption         =   "Prov. Num2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   4800
      Width           =   840
   End
   Begin VB.Label lblProvNum1 
      AutoSize        =   -1  'True
      Caption         =   "Prov. Num1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   4440
      Width           =   840
   End
   Begin VB.Label lblCountry 
      AutoSize        =   -1  'True
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   540
   End
   Begin VB.Label lblProvNo 
      AutoSize        =   -1  'True
      Caption         =   "Prov. #"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   525
   End
   Begin VB.Label lblProvName 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label lblProvCode 
      AutoSize        =   -1  'True
      Caption         =   "Code"
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   450
   End
End
Attribute VB_Name = "frmPROV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew As Integer
Dim fglbNewRec%, xOldCode As String

Private Function chkProv()
Dim prov$, SQLQ$, Msg$
Dim snapProv As New ADODB.Recordset

chkProv = False

On Error GoTo chkProv_Err

If Len(txtCode) < 1 Then
    MsgBox "Province/State is a required field"
    txtCode.SetFocus
    Exit Function
End If

If Len(txtDesc) < 1 Then
    MsgBox "Description of Province/State is a required field"
    txtDesc.SetFocus
    Exit Function
End If

If Len(txtCountry) < 1 Then
    MsgBox "Country of Province/State is a required field"
    txtCountry.SetFocus
    Exit Function
End If

'Ticket #13557 Frank Aug 22, 2007
If Not UCase(txtCountry) = "AUSTRALIA" Then
    If Len(txtCode) > 2 Then
        MsgBox "The maximum length for Province/State code is two characters" & Chr(10) & " except for Australia which has three characters"
        txtCode.SetFocus
        Exit Function
    End If
End If


prov$ = txtCode

If fglbNewRec Then
    SQLQ$ = "SELECT * from HRPROV "
    SQLQ$ = SQLQ$ & " WHERE HRPROV.CODE = '" & prov$ & "'"
    SQLQ$ = SQLQ$ & " ORDER BY HRPROV.CODE "
    snapProv.Open SQLQ$, gdbAdoIhr001, adOpenStatic
    If snapProv.BOF And snapProv.EOF Then
        snapProv.Close
    Else
        Msg$ = "This Province/State Code already exists"
        MsgBox Msg$
        snapProv.Close
        txtCode.SetFocus
        Exit Function
    End If
End If

'Release 8.0 - Ticket #22682: New fields
If Len(txtProvNum1.Text) > 0 Then
    If Not IsNumeric(txtProvNum1.Text) Then
        MsgBox "Invalid " & lStr(lblProvNum1.Caption) & ". It must be numeric."
        txtProvNum1.SetFocus
        Exit Function
    End If
End If

If Len(txtProvNum2.Text) > 0 Then
    If Not IsNumeric(txtProvNum2.Text) Then
        MsgBox "Invalid " & lStr(lblProvNum2.Caption) & ". It must be numeric."
        txtProvNum2.SetFocus
        Exit Function
    End If
End If

SQLQ$ = "SELECT * from HRPROV "
SQLQ$ = SQLQ$ & " WHERE HRPROV.CODE = '" & prov$ & "'"
'SQLQ$ = SQLQ$ & " AND UPPER(HRPROV.COUNTRY) = '" & UCase(txtCountry.Text) & "'"
SQLQ$ = SQLQ$ & " AND (HRPROV.COUNTRY) = '" & (txtCountry.Text) & "'"
If Not fglbNewRec Then
    SQLQ$ = SQLQ$ & " AND HRPROV.CODE <> '" & xOldCode & "'"
End If
SQLQ$ = SQLQ$ & " ORDER BY HRPROV.CODE "

snapProv.Open SQLQ$, gdbAdoIhr001, adOpenStatic
If snapProv.BOF And snapProv.EOF Then
    snapProv.Close
Else
    Msg$ = "This Province/State already exists"
    MsgBox Msg$
    snapProv.Close
    Exit Function
End If

chkProv = True

Exit Function

chkProv_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Select", "HRPROv", "Cancel")
Resume Next

End Function



Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

'Data1.UpdateControls    ' returns without saving
Data1.Recordset.CancelBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Call modSTUPD(False)  ' reset screen's attributes


cmdClose.SetFocus

fglbNewRec% = False

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
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Msg As String, a%

On Error GoTo DelErr

If Data1.Recordset.RecordCount <= 1 Then
    MsgBox "You Can Not Delete The Last Province."
    Exit Sub
End If

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Province?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then
    Exit Sub
End If

Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh



Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Single", "Delete")
Call RollBack '09June99 js

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim SQLQ$

txtFindKey.SetFocus 'added by Marlon Cowan 9/16/97

If Len(txtFindKey) > 0 Then
    SQLQ$ = "CODE >= '" & txtFindKey.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ$
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    Exit Sub
End If

If Len(txtFindDesc) > 0 Then
    SQLQ$ = "DESCR >= '" & txtFindDesc.Text & "'"
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ$
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

Call modSTUPD(True)
txtCode.Enabled = True
txtCode.SetFocus
xOldCode = txtCode

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '09June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

On Error GoTo NewErr


glbCodeRef = True


Data1.Recordset.AddNew
txtComp.Text = glbCompNo
fglbNewRec% = True
Call modSTUPD(True)
txtCode.SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HRPROV", "AddNew")
Resume Next

End Sub

Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
'Dim SQLQ As String
Dim Desc As String
Dim ProvCode
On Error GoTo OK_Err

If Not chkProv() Then Exit Sub
ProvCode = txtCode

Data1.Recordset("CODE") = txtCode & ""
Data1.Recordset("COUNTRY") = txtCountry & ""

'Release 8.0 - Ticket #22682: Add new fields
Data1.Recordset("PR_NUM1") = IIf(txtProvNum1.Text = "", 0, txtProvNum1.Text) & ""
Data1.Recordset("PR_NUM2") = IIf(txtProvNum2.Text = "", 0, txtProvNum2.Text) & ""
Data1.Recordset("PR_TXT1") = txtProvText1.Text & ""
Data1.Recordset("PR_TXT2") = txtProvText2.Text & ""
Data1.Recordset("PR_TXT3") = txtProvText3.Text & ""
Data1.Recordset.UpdateBatch

If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Data1.Recordset.Find "CODE='" & ProvCode & "'"

fglbNewRec% = False


Call modSTUPD(False)

cmdClose.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption

Resume Next
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

RHeading = "Province/State Codes"
Me.vbxCrystal.WindowTitle = "Province/State Code Report"
Me.vbxCrystal.BoundReportHeading = "Province/State Codes"
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdSelect_Click()

If Me.vbxTrueGrid.EditActive = True Then
    MsgBox "Save/Cancel changes first"
Else
    glbProv = Data1.Recordset("CODE")
    glbProvDesc = Data1.Recordset("DESCR")
    Unload frmPROV
End If

End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRPROV", "SELECT")

End Sub

Private Sub Form_Load()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT * FROM HRPROV"
Data1.Refresh

lblProvName.Caption = lStr("Prov. Name")
lblProvNo.Caption = lStr("Prov. #")
lblProvNum1.Caption = lStr("Prov. Num1")
lblProvNum2.Caption = lStr("Prov. Num2")
lblProvTxt1.Caption = lStr("Prov. Text1")
lblProvTxt2.Caption = lStr("Prov. Text2")
lblProvTxt3.Caption = lStr("Prov. Text3")

vbxTrueGrid.Columns(1).Caption = lStr("Prov. Name")
vbxTrueGrid.Columns(2).Caption = lStr("Prov. #")
vbxTrueGrid.Columns(4).Caption = lStr("Prov. Num1")
vbxTrueGrid.Columns(5).Caption = lStr("Prov. Num2")
vbxTrueGrid.Columns(6).Caption = lStr("Prov. Text1")
vbxTrueGrid.Columns(7).Caption = lStr("Prov. Text2")
vbxTrueGrid.Columns(8).Caption = lStr("Prov. Text3")

Call modSTUPD(False)            'Jaddy 10/18/99

Screen.MousePointer = DEFAULT   '
                                
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

cmdOK.Enabled = TF          'May99 js
cmdCancel.Enabled = TF      '
cmdClose.Enabled = FT
cmdPrint.Enabled = FT       '
cmdFind.Enabled = FT        '
cmdSelect.Enabled = FT
If gSec_Province Then '
    cmdModify.Enabled = FT      '
    cmdNew.Enabled = FT         '
    cmdDelete.Enabled = FT      '
Else
    cmdModify.Enabled = False      '
    cmdNew.Enabled = False   '
    cmdDelete.Enabled = False
End If

txtCode.Enabled = TF        '
txtDesc.Enabled = TF        '
txtNumber.Enabled = TF
txtCountry.Enabled = TF

'Release 8.0 - Ticket #22682: Add new fields
txtProvNum1.Enabled = TF
txtProvNum2.Enabled = TF
txtProvText1.Enabled = TF
txtProvText2.Enabled = TF
txtProvText3.Enabled = TF

txtFindDesc.Enabled = FT
txtFindKey.Enabled = FT
vbxTrueGrid.Enabled = FT

If glbDivInhSel Then
    cmdSelect.Enabled = False
End If

End Sub

Private Sub txtCode_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCountry_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCountry_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtDesc_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtNumber_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub txtProvNum1_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtProvNum1_KeyPress(KeyAscii As Integer)
If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtProvNum1_LostFocus()
If Not IsNumeric(Trim(txtProvNum1)) And txtProvNum1.DataChanged Then txtProvNum1 = 0
End Sub

Private Sub txtProvNum2_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtProvNum2_KeyPress(KeyAscii As Integer)
If Not IsNumericEntry(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtProvNum2_LostFocus()
If Not IsNumeric(Trim(txtProvNum2)) And txtProvNum2.DataChanged Then txtProvNum2 = 0
End Sub

Private Sub txtProvText1_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtProvText2_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtProvText3_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_DblClick()

If Not Me.vbxTrueGrid.EditActive Then
    glbProv = Data1.Recordset("CODE")
    glbProvDesc = Data1.Recordset("DESCR")
    Unload frmPROV
Else
    MsgBox "Save/cancel changes first"
End If

End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HRPROV"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh

End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    glbProv = Data1.Recordset("CODE")
    glbProvDesc = Data1.Recordset("DESCR")
    If Me.vbxTrueGrid.EditActive Then  '.EditActive
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

Private Function IsNumericEntry(KeyAscii As Integer, Optional NegAllowed As Boolean) As Boolean
    If KeyAscii = Asc(vbBack) Or IsNumeric(Chr(KeyAscii)) Or (NegAllowed And KeyAscii = Asc("-")) Then IsNumericEntry = True
End Function

