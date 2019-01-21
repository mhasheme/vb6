VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEFOLLOWUP 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Follow-Up Data"
   ClientHeight    =   8130
   ClientLeft      =   285
   ClientTop       =   1350
   ClientWidth     =   10230
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
   ScaleHeight     =   8130
   ScaleWidth      =   10230
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fefollow.frx":0000
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "fefollow.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   9015
   End
   Begin INFOHR_Controls.DateLookup dlpEDate 
      DataField       =   "EF_FDATE"
      Height          =   285
      Left            =   6690
      TabIndex        =   2
      Tag             =   "41-Effective Date of Followup"
      Top             =   2520
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EF_ADMINBY"
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Tag             =   "00-Enter Administered By Code"
      Top             =   2880
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EF_FREAS"
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Tag             =   "01-Followup Reason Code"
      Top             =   2520
      Width           =   3885
      _ExtentX        =   6853
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "FURE"
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   375
      Left            =   8520
      Top             =   7080
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Ado2"
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
      TabIndex        =   20
      Top             =   7470
      Width           =   10230
      _Version        =   65536
      _ExtentX        =   18045
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
      Begin VB.CommandButton cmdMarkAll 
         Appearance      =   0  'Flat
         Caption         =   "&Mark All Comp."
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Tag             =   "Mark all messages as being completed."
         Top             =   0
         Width           =   1620
      End
      Begin VB.CommandButton cmdMassDelete 
         Appearance      =   0  'Flat
         Caption         =   "Ma&ss Delete"
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Tag             =   "Remove all messages listed above."
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1770
         TabIndex        =   21
         Tag             =   "Save the changes made"
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9120
         Top             =   240
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
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EF_LDATE"
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
      DataField       =   "EF_LTIME"
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
      Left            =   4440
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5880
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EF_LUSER"
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
      Width           =   10230
      _Version        =   65536
      _ExtentX        =   18045
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
         Left            =   7560
         TabIndex        =   24
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
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
         Left            =   1575
         TabIndex        =   11
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
         Left            =   3360
         TabIndex        =   10
         Top             =   135
         Width           =   720
      End
   End
   Begin Threed.SSCheck chkCompleted 
      DataField       =   "EF_COMPLETED"
      Height          =   240
      Left            =   5640
      TabIndex        =   4
      Tag             =   "00-Followup Completed"
      Top             =   2850
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   423
      _StockProps     =   78
      Caption         =   "Completed    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "EF_COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Tag             =   "00-Comments - free form"
      Text            =   "fefollow.frx":355C
      Top             =   3540
      Width           =   8805
   End
   Begin VB.Label lblCompleted 
      Height          =   255
      Left            =   7770
      TabIndex        =   19
      Top             =   2880
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblAdminBy 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
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
      Left            =   360
      TabIndex        =   18
      Top             =   2880
      Width           =   1125
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   17
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   5700
      TabIndex        =   16
      Top             =   2520
      Width           =   780
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reason"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   15
      Top             =   2520
      Width           =   660
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EF_EMPNBR"
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
      DataField       =   "EF_COMPNO"
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
Attribute VB_Name = "frmEFOLLOWUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim rsGrid As ADODB.Recordset
Dim oEffDate
'Dim FRS As ADODB.Recordset

Private Sub chkCompleted_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

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
    MsgBox "Reason code is a required field"
    clpCode(1).SetFocus
    Exit Function
End If


If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Reason code must be valid"
    clpCode(1).SetFocus
    Exit Function
Else
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    'SQLQ = "SELECT ACCESSABLE from HR_SECURE_FOLLOW_UP WHERE USERID='" & glbUserID & "'"
    SQLQ = SQLQ & " AND CODENAME='" & clpCode(1).Text & "'"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        If rs("MAINTAINABLE") = 0 Then
        'If rs("ACCESSABLE") = 0 Then
            MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(1).Text & "' Reason Code.", vbOKOnly + vbInformation, "Authorization failed"
            rs.Close
            Set rs = Nothing
            clpCode(1).SetFocus
            Exit Function
        End If
    Else
        MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(1).Text & "' Reason Code.", vbOKOnly + vbInformation, "Authorization failed"
        rs.Close
        Set rs = Nothing
        clpCode(1).SetFocus
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
End If

If clpCode(2).Caption = "Unassigned" And Len(Trim(clpCode(2).Text)) > 0 Then
    MsgBox lStr("Administered By") & " type must be valid"
    clpCode(2).SetFocus
    Exit Function
End If

If Len(dlpEDate.Text) >= 1 Then
    If Not IsDate(dlpEDate.Text) Then
        MsgBox "Effective Date is not a valid date."
        dlpEDate.SetFocus
        Exit Function
    End If
Else
    MsgBox "Effective Date is required."
    dlpEDate.SetFocus
    Exit Function
End If

If chkCompleted.Value = True Then
    lblCompleted.Caption = "Y"
Else
    lblCompleted.Caption = "N"
End If

chkEComment = True

Exit Function

chkEComment_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkEFollow", "HR_EFOLLOWUP", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Sub cmdCancel_Click()

On Error GoTo Can_Err

'Data2.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data2.Refresh
''' Sam add July 2002 * Remove Binding Control
rsDATA.CancelUpdate

fglbNew = False
Call SET_UP_MODE
Call Display_Value
Data1.Refresh

'Call ST_UPD_MODE(True)  ' reset screen's attributes

'Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_FOLLOW_UP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEFOLLOWUP" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String


If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

If Not FollowUp_Sec Then
    MsgBox "You do not have Authority to complete this Reason Code transaction.", vbInformation + vbOKOnly, "Authorization failure"
    Exit Sub
End If


On Error GoTo Del_Err

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

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_FOLLOW_UP", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub cmdDelete_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdMarkAll_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant

On Error GoTo MAll_Err

Msg$ = lStr("Mark all Follow-Up Records")
Msg = Msg$ & Chr(10) & "listed as completed?"
Title$ = "Mark all completed?"   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    'Friesens - Ticket #16591
    If glbCompSerial = "S/N - 2279W" Then
        gdbAdoIhr001.Execute "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "' WHERE EF_EMPNBR=" & glbLEE_ID & " AND (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1)"
    Else
        gdbAdoIhr001.Execute "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = 1, EF_LDATE = " & Date_SQL(Date) & ", EF_LTIME = '" & Time$ & "', EF_LUSER = '" & glbUserID & "' WHERE EF_EMPNBR=" & glbLEE_ID & " AND EF_COMPLETED <> 1"
    End If
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
End If


Exit Sub

MAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdMarkAll", "HR_FOLLOW_UP", "Mark All")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdMarkAll_GotFocus()
Call SetPanHelp(ActiveControl)

End Sub

Private Sub cmdMassDelete_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant
On Error GoTo MDel_Err


Msg$ = lStr("Delete all Follow-Up Records")
Msg = Msg$ & Chr(10) & "listed?"
Title$ = lStr("Delete all Follow-Up Records")   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    'Friesens - Ticket #16591
    If glbCompSerial = "S/N - 2279W" Then
        gdbAdoIhr001.Execute "DELETE FROM HR_FOLLOW_UP WHERE EF_EMPNBR=" & glbLEE_ID & " AND (EF_FREAS <> 'EDUC' or EF_COMPLETED = 1)"
    Else
        gdbAdoIhr001.Execute "DELETE FROM HR_FOLLOW_UP WHERE EF_EMPNBR=" & glbLEE_ID
    End If
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    Screen.MousePointer = DEFAULT
    Unload Me
End If

Exit Sub


MDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_FOLLOW_UP", "Delete")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdMassDelete_GotFocus()
Call SetPanHelp(ActiveControl)

End Sub

'Private Sub cmdModify_GotFocus()
'Call SetPanHelp(ActiveControl)
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
'lblEEID = glbLEE_ID

lblCNum.Caption = "001"

clpCode(1).SetFocus
chkCompleted.Value = False

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_FOLLOW_UP", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Private Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()

On Error GoTo Add_Err

If Not chkEComment() Then Exit Sub

'Ticket #22682 - Release 8.0: Follow Up Email Sending
If fglbNew = False Then
    If oEffDate <> dlpEDate Then
        rsDATA("EF_EMAIL_SENT") = "N"
        rsDATA("EF_EMAIL_DATE") = Null
    End If
End If

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

'Ticket #22682 - Release 8.0: Follow Up Email Sending
oEffDate = dlpEDate

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

fglbNew = False

Call ST_UPD_MODE(True)
Call SET_UP_MODE

'Me.vbxTrueGrid.SetFocus
If NextFormIF("Follow-Up") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    'Data1.Recordset.CancelUpdate
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_FOLLOW_UP", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String ', glbstrSelCri
    
'    cmdPrint.Enabled = False
    RHeading = lblEEName & lStr("'s Follow-ups")
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    
    If Not glbtermopen Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfollup.rpt"
        
        glbstrSelCri = "{HR_FOLLOW_UP.EF_EMPNBR} = " & glbLEE_ID & " "
        If (glbNoNONE And glbUNION = "NONE") Or (glbNoEXEC And glbUNION = "EXEC") Then      'Hemu -EXE
            glbstrSelCri = glbstrSelCri & " And {HR_FOLLOW_UP.EF_FREAS} <> 'SREV' "
        End If
        
        Call Cri_Sec
        
        'Follow Up security
        'glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP.ACCESSABLE} = True"
        
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
        
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRDB
            Me.vbxCrystal.DataFiles(1) = glbIHRDB
        End If
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfollupT.rpt"
        
        glbstrSelCri = "{Term_FOLLOW_UP.TERM_SEQ}=" & glbTERM_Seq & " "
        If (glbNoNONE And glbUNION = "NONE") Or (glbNoEXEC And glbUNION = "EXEC") Then      'Hemu -EXE
            glbstrSelCri = glbstrSelCri & " And {Term_FOLLOW_UP.EF_FREAS} <> 'SREV' "
        End If
        
        Call Cri_Sec
        
        'Follow Up security
        'glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP.ACCESSABLE} = True"
        
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
        
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
'    cmdPrint.Enabled = True
End Sub

Sub cmdView_Click()
Dim RHeading As String ', glbstrSelCri
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
'    cmdPrint.Enabled = False
    RHeading = lblEEName & lStr("'s Follow-ups")
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
        
    If Not glbtermopen Then
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfollup.rpt"
        
        glbstrSelCri = "{HR_FOLLOW_UP.EF_EMPNBR} = " & glbLEE_ID & " "
        If (glbNoNONE And glbUNION = "NONE") Or (glbNoEXEC And glbUNION = "EXEC") Then      'Hemu -EXE
            glbstrSelCri = glbstrSelCri & " And {HR_FOLLOW_UP.EF_FREAS} <> 'SREV' "
        End If
        
        Call Cri_Sec
        
        'Follow Up security
        'glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP.ACCESSABLE} = True"
        
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
        
        If glbSQL Or glbOracle Then
            Me.vbxCrystal.Connect = RptODBC_SQL
        Else
            Me.vbxCrystal.Connect = "PWD=petman;"
            Me.vbxCrystal.DataFiles(0) = glbIHRDB
            Me.vbxCrystal.DataFiles(1) = glbIHRDB
        End If
    Else
        Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfollupT.rpt"
        
        glbstrSelCri = "{Term_FOLLOW_UP.TERM_SEQ}=" & glbTERM_Seq & " "
        If glbNoNONE And glbUNION = "NONE" Or (glbNoEXEC And glbUNION = "EXEC") Then
            glbstrSelCri = glbstrSelCri & " And {HR_FOLLOW_UP.EF_FREAS} <> 'SREV' "
        End If
        
        Call Cri_Sec
        
        'Follow Up security
        'glbstrSelCri = glbstrSelCri & " AND {HR_SECURE_FOLLOW_UP.USERID} ='" & glbUserID & "' AND {HR_SECURE_FOLLOW_UP.ACCESSABLE} = True"
        
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
        
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
'    cmdPrint.Enabled = True
End Sub

'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
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

' created query below for multiple joins - might as well use it.
' as left join not editable

 'Release 8.0 - Ticket #22682: Get Employee # of the User - View Own security
If Not glbtermopen Then
    If glbUserEmpNo = glbLEE_ID And Not gSec_FollUp_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("Follow-ups") & " information.", vbCritical, "info:HR - Security"
        'glbLEE_ID = 0      'Ticket #25208
        Screen.MousePointer = DEFAULT
        Unload Me: Exit Function
    End If
End If

 ' out or left join query not updateable - so do straight.
If glbtermopen Then
    SQLQ = " "
    'SQLQ = SQLQ & "SELECT Term_FOLLOW_UP.* "
    'SQLQ = SQLQ & "FROM Term_FOLLOW_UP "
    'SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    
    SQLQ = SQLQ & "SELECT Term_FOLLOW_UP.* FROM Term_FOLLOW_UP INNER JOIN HR_SECURE_FOLLOW_UP ON Term_FOLLOW_UP.EF_FREAS = HR_SECURE_FOLLOW_UP.CODENAME"
    'SQLQ = SQLQ & "FROM Term_FOLLOW_UP "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
    End If
Else
    SQLQ = " "
    'SQLQ = SQLQ & "SELECT HR_FOLLOW_UP.* "
    'SQLQ = SQLQ & "FROM HR_FOLLOW_UP "
    'SQLQ = SQLQ & "WHERE (HR_FOLLOW_UP.EF_EMPNBR = " & glbLEE_ID & ") "
    
    SQLQ = "SELECT HR_FOLLOW_UP.* FROM HR_FOLLOW_UP INNER JOIN HR_SECURE_FOLLOW_UP ON HR_FOLLOW_UP.EF_FREAS = HR_SECURE_FOLLOW_UP.CODENAME"
    SQLQ = SQLQ & " WHERE HR_FOLLOW_UP.EF_EMPNBR = " & glbLEE_ID
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
    End If
End If

If (glbNoNONE And glbUNION = "NONE") Or (glbNoEXEC And glbUNION = "EXEC") Then     'Hemu -EXE
    SQLQ = SQLQ & " and EF_FREAS <> 'SREV' "
End If

'Hemu
SQLQ = SQLQ & " ORDER BY EF_FDATE DESC"     'Ticket #28635 - DESC order
'Hemu

Data1.RecordSource = SQLQ
Data1.Refresh

Set rsGrid = Data1.Recordset.Clone
vbxTrueGrid.FetchRowStyle = True

EERetrieve = True
Screen.MousePointer = DEFAULT


Exit Function


EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SklsRetrieve", "HR_FOLLOW_UP", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
Exit Function
End Function

Private Sub Form_Activate()
glbOnTop = "FRMEFOLLOWUP"
Call SET_UP_MODE

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    frmEEFIND.Show 1
Else
    Me.Show
    lblEEID = glbLEE_ID
End If

End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMEFOLLOWUP"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEFOLLOWUP"

If glbtermopen Then
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
    If glbUserEmpNo = glbLEE_ID And Not gSec_FollUp_ViewOwn Then
        MsgBox "You cannot view your own " & lStr("Follow-ups") & " information.", vbCritical, "info:HR - Security"
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

'If Len(glbLEE_SName) < 1 Then Exit Sub
If Len(glbLEE_SName) < 1 Then 'Exit Sub
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
End If

Screen.MousePointer = HOURGLASS

Me.vbxTrueGrid.SetFocus

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = lStr("Follow-ups - ") & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

lblEENum.Caption = ShowEmpnbr(lblEEID) 'glbLEE_ID

Call setCaption(lblAdminBy)
Call ST_UPD_MODE(False)

'vbxTrueGrid.FetchRowStyle = True
'vbxTrueGrid.MarqueeStyle = 3

Call Display_Value

If Not gSec_Upd_Follow_Ups Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
    cmdMarkAll.Enabled = False
    cmdMassDelete.Enabled = False
End If

Call INI_Controls(Me)
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False



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
Set frmEFOLLOWUP = Nothing
Call NextForm
End Sub

Private Sub frmDetails_Click()

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


fUPMode = TF    ' update mode

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT

'cmdCancel.Enabled = TF
'cmdOK.Enabled = TF
chkCompleted.Enabled = TF
memComments.Enabled = TF
dlpEDate.Enabled = TF

cmdMarkAll.Enabled = TF 'FT
cmdMassDelete.Enabled = TF 'FT
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
vbxTrueGrid.Enabled = False
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
vbxTrueGrid.Enabled = True

End Sub

Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
On Error GoTo Eh
'Friesens - Ticket #16591
If glbCompSerial = "S/N - 2279W" Then
    rsGrid.Bookmark = Bookmark
    If Not IsNull(rsGrid("EF_FREAS")) And rsGrid("EF_FREAS") <> "" Then
        'Disable changes to EDUC records only
        If rsGrid("EF_FREAS") = "EDUC" And rsGrid("EF_COMPLETED") = False Then
            'Grey the text
            RowStyle.ForeColor = vbGrayText
            
            'Disable the controls - to prevent changes to the records
            clpCode(1).Enabled = False
            clpCode(2).Enabled = False
            dlpEDate.Enabled = False
            chkCompleted.Enabled = False
            memComments.Enabled = False
        Else
            'Enable controls
            clpCode(1).Enabled = True
            clpCode(2).Enabled = True
            dlpEDate.Enabled = True
            chkCompleted.Enabled = True
            memComments.Enabled = True
        End If
    End If
End If
Eh:
    Exit Sub
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
    
    If glbtermopen Then
        'SQLQ = SQLQ & "SELECT Term_FOLLOW_UP.* "
        'SQLQ = SQLQ & "FROM Term_FOLLOW_UP "
        'SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    
        SQLQ = SQLQ & "SELECT Term_FOLLOW_UP.* FROM Term_FOLLOW_UP INNER JOIN HR_SECURE_FOLLOW_UP ON Term_FOLLOW_UP.EF_FREAS = HR_SECURE_FOLLOW_UP.CODENAME"
        'SQLQ = SQLQ & "FROM Term_FOLLOW_UP "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        If xTemplate = "" Or xTemplate = "TEMPLATE" Then
            SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
        Else
            '????Ticket #24808 -  Retrieve template's security profile
            SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
        End If
    Else
        'SQLQ = SQLQ & "SELECT HR_FOLLOW_UP.* "
        'SQLQ = SQLQ & "FROM HR_FOLLOW_UP "
        'SQLQ = SQLQ & "WHERE (HR_FOLLOW_UP.EF_EMPNBR = " & glbLEE_ID & ") "
        
        SQLQ = "SELECT HR_FOLLOW_UP.* FROM HR_FOLLOW_UP INNER JOIN HR_SECURE_FOLLOW_UP ON HR_FOLLOW_UP.EF_FREAS = HR_SECURE_FOLLOW_UP.CODENAME"
        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
        If xTemplate = "" Or xTemplate = "TEMPLATE" Then
            SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(glbUserID, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
        Else
            '????Ticket #24808 -  Retrieve template's security profile
            SQLQ = SQLQ & " AND HR_SECURE_FOLLOW_UP.USERID ='" & Replace(xTemplate, "'", "''") & "' AND HR_SECURE_FOLLOW_UP.ACCESSABLE <> 0"
        End If
    End If
    
    If (glbNoNONE And glbUNION = "NONE") Or (glbNoEXEC And glbUNION = "EXEC") Then     'Hemu -EXE
        SQLQ = SQLQ & " and EF_FREAS <> 'SREV' "
    End If

    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
  
    Data1.RecordSource = SQLQ
    Data1.Refresh
        
    Set rsGrid = Data1.Recordset.Clone
    vbxTrueGrid.FetchRowStyle = True

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
Dim SQLQ As String, x%
Dim ButtonS As New ButtonsSetting

On Error GoTo Tab1_Err
Call Display_Value

'Ticket #22682 - Release 8.0: Follow Up Email Sending
oEffDate = dlpEDate

'Friesens - Ticket #16591
If glbCompSerial = "S/N - 2279W" Then
    If Not IsNull(clpCode(1).Text) And clpCode(1).Text <> "" Then
        'Disable changes to EDUC records only
        If clpCode(1).Text = "EDUC" And chkCompleted.Value = False Then
            'Disable the controls - to prevent changes to the records
            clpCode(1).Enabled = False
            clpCode(2).Enabled = False
            dlpEDate.Enabled = False
            chkCompleted.Enabled = False
            memComments.Enabled = False
            
            ButtonS.Enabled("delete") = False
            ButtonS.Enabled("save") = False
        Else
            'Enable controls
            clpCode(1).Enabled = True
            clpCode(2).Enabled = True
            dlpEDate.Enabled = True
            chkCompleted.Enabled = True
            memComments.Enabled = True
            
            ButtonS.Enabled("delete") = True
            ButtonS.Enabled("save") = True
        End If
    End If
End If

Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_EFOLLOWUP", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub chkCompleted_Click(Value As Integer)
    
If chkCompleted.Value = True Then
    lblCompleted.Caption = "Y"
Else
    lblCompleted.Caption = "N"
End If

End Sub

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
            SQLQ = " "
            SQLQ = SQLQ & "SELECT Term_FOLLOW_UP.* "
            SQLQ = SQLQ & "FROM Term_FOLLOW_UP "
            SQLQ = SQLQ & "WHERE (EF_FOLLOWUP_ID= " & Data1.Recordset!EF_FOLLOWUP_ID & ") "
            
            If glbNoNONE And glbUNION = "NONE" Then
                SQLQ = SQLQ & " and EF_FREAS <> 'SREV' "
            End If
            
            'Hemu
            SQLQ = SQLQ & " ORDER BY EF_FDATE DESC"
            'Hemu
            
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            SQLQ = " "
            SQLQ = SQLQ & "SELECT HR_FOLLOW_UP.* "
            SQLQ = SQLQ & "FROM HR_FOLLOW_UP "
            SQLQ = SQLQ & "WHERE (EF_FOLLOWUP_ID= " & Data1.Recordset!EF_FOLLOWUP_ID & ") "
        
            If glbNoNONE And glbUNION = "NONE" Then
                SQLQ = SQLQ & " and EF_FREAS <> 'SREV' "
            End If
            
            'Hemu
            SQLQ = SQLQ & " ORDER BY EF_FDATE DESC"     'Ticket #28635 - DESC order
            'Hemu
            
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
UpdateRight = gSec_Upd_Follow_Ups
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
    TF = FollowUp_Sec()
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    'frmVATTEND.Caption = "Attendance - " & Left$(glbLEE_SName, 5)
    frmEFOLLOWUP.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Function FollowUp_Sec() As Boolean
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim retVal As Boolean
    Dim xTemplate As String
    
    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbUserID)
    
    
    strSQL = "SELECT MAINTAINABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    'strSQL = "SELECT ACCESSABLE FROM HR_SECURE_FOLLOW_UP WHERE "
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(glbUserID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        strSQL = strSQL & "CODENAME='" & clpCode(1).Text & "' AND USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF = False And rs.BOF = False Then
        retVal = Abs(rs("MAINTAINABLE"))
        'retVal = Abs(rs("ACCESSABLE"))
    Else
        retVal = False
    End If
    
    FollowUp_Sec = retVal
End Function

Private Sub Cri_Sec()
    Dim EECri As String
    Dim strSec As String
    
    strSec = buildSec_FollowUp
    If Len(strSec) >= 1 Then
        If Not glbtermopen Then
            EECri = "{HR_FOLLOW_UP.EF_FREAS} " & Replace(Replace(strSec, "(", "["), ")", "]")
        Else
            EECri = "{Term_FOLLOW_UP.EF_FREAS} " & Replace(Replace(strSec, "(", "["), ")", "]")
        End If
    End If
    
    If Len(EECri) >= 1 Then
        If Len(glbstrSelCri) > 0 Then
            glbstrSelCri = glbstrSelCri & " AND " & EECri
        Else
            glbstrSelCri = EECri
        End If
    End If
    
End Sub


