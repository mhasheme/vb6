VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSReOccur 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Reoccurrence"
   ClientHeight    =   8490
   ClientLeft      =   150
   ClientTop       =   180
   ClientWidth     =   11580
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
   ScaleHeight     =   8490
   ScaleWidth      =   11580
   WindowState     =   2  'Maximized
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "CC_COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2460
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "00-Comments"
      Top             =   4560
      Width           =   7035
   End
   Begin INFOHR_Controls.DateLookup dlpReturnDate 
      DataField       =   "CC_RETURNDATE"
      Height          =   285
      Left            =   2145
      TabIndex        =   4
      Tag             =   "42-Cost To Date"
      Top             =   3840
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpApprDate 
      DataField       =   "CC_APPROVDATE"
      Height          =   285
      Left            =   2145
      TabIndex        =   3
      Tag             =   "42-Cost From Date"
      Top             =   3480
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpReoccur 
      DataField       =   "CC_REOCCURDATE"
      Height          =   285
      Left            =   2130
      TabIndex        =   2
      Tag             =   "41-WSIB Cost Statement Date"
      Top             =   3120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   10320
      Top             =   8040
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      TabIndex        =   0
      Top             =   7830
      Width           =   11580
      _Version        =   65536
      _ExtentX        =   20426
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
         Left            =   6840
         Top             =   240
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
   Begin VB.ComboBox cmbWCB 
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
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2460
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Choose WSIB this cost is related to"
      Top             =   2700
      Width           =   2655
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CC_LDATE"
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
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CC_LTIME"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "CC_LUSER"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   7020
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11580
      _Version        =   65536
      _ExtentX        =   20426
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
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   135
         Width           =   720
      End
   End
   Begin MSMask.MaskEdBox medShifts 
      DataField       =   "CC_SHIFTSLOST"
      Height          =   285
      Left            =   2460
      TabIndex        =   5
      Tag             =   "20-Cost for Temporary Compensation"
      Top             =   4200
      Width           =   1485
      _ExtentX        =   2619
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
      Format          =   "0"
      PromptChar      =   "_"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fehsreoc.frx":0000
      Height          =   1935
      Left            =   240
      OleObjectBlob   =   "fehsreoc.frx":0014
      TabIndex        =   20
      Top             =   600
      Width           =   9015
   End
   Begin VB.Label lblClaimData 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "CC_WCBNBR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   23
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbComment 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      Left            =   240
      TabIndex        =   22
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label lblClaim 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Claim #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   2760
      Width           =   1245
   End
   Begin VB.Label lblReturn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Return To Work Date"
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
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   1545
   End
   Begin VB.Label lblApprDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Approved"
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
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   1080
   End
   Begin VB.Label lblShifts 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Shifts Lost"
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
      Left            =   210
      TabIndex        =   17
      Top             =   4200
      Width           =   735
   End
   Begin VB.Label lblReoccur 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Reoccurrence"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   210
      TabIndex        =   16
      Top             =   3135
      Width           =   1905
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CC_EMPNBR"
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
      TabIndex        =   14
      Top             =   7140
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CC_COMPNO"
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
      Left            =   315
      TabIndex        =   15
      Top             =   7140
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEHSReOccur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim x%
Dim fglbNew
Dim wcb() As Variant
Dim fglbComboWCB% ' is there data in combo box? were wcbs found?
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Private Function chkHSWCBCs()
Dim SQLQ As String, Msg As String, dd&, tdat As Variant
Dim rsTemp As New ADODB.Recordset
chkHSWCBCs = False

On Error GoTo chkHSWCBCs_Err
If Len(cmbWCB.Text) = 0 Then
    MsgBox "Claim# is required."
    dlpReoccur.SetFocus
    Exit Function
End If
'Tested this function with Linda, she said the duplicate records were OK.
'If fglbNew Then 'Check duplicate record
'    If glbtermopen Then
'        SQLQ = "SELECT * from Term_OHS_REOCCURENCE "
'        SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq & " "
'        SQLQ = SQLQ & "AND CC_WCBNBR ='" & cmbWCB.Text & "' "
'        If rsTemp.State <> 0 Then rsTemp.Close
'        rsTemp.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
'    Else
'        SQLQ = "SELECT * from HR_OHS_REOCCURENCE "
'        SQLQ = SQLQ & "WHERE CC_EMPNBR = " & glbLEE_ID & " "
'        SQLQ = SQLQ & "AND CC_WCBNBR ='" & cmbWCB.Text & "' "
'        If rsTemp.State <> 0 Then rsTemp.Close
'        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    End If
'    If Not rsTemp.EOF Then
'        rsTemp.Close
'        MsgBox "Duplicate record found"
'        dlpReoccur.SetFocus
'        Exit Function
'    End If
'    rsTemp.Close
'End If

If Len(dlpReoccur.Text) >= 1 Then
    If Not IsDate(dlpReoccur.Text) Then
        MsgBox "Date of Reoccurrence is not a valid date."
        dlpReoccur.SetFocus
        Exit Function
    End If
Else
    MsgBox "Date of Reoccurrence is required."
    dlpReoccur.SetFocus
    Exit Function
End If

If Len(dlpApprDate.Text) >= 1 Then
    If Not IsDate(dlpApprDate.Text) Then
        MsgBox "Date Approved is not a valid date."
        dlpApprDate.SetFocus
        Exit Function
    End If
End If

If Len(dlpReturnDate.Text) >= 1 Then
    If Not IsDate(dlpReturnDate.Text) Then
        MsgBox "Return Work Date is not a valid date."
        dlpReturnDate.SetFocus
        Exit Function
    End If
End If


'tdat = wcb(cmbWCB.ListIndex + 1, 3)

Dim Ctrol As Control

Set Ctrol = medShifts: If Not chkNumeric(Ctrol) Then Exit Function


chkHSWCBCs = True

Exit Function

chkHSWCBCs_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_OHS_REOCCURENCE", "HR_OHS_REOCCURENCE", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function



Private Sub cmbWCB_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmbWCB_LostFocus()
lblClaimData.Caption = cmbWCB.Text
'Dim X%
'X% = cmbWCB.ListIndex
'If X% >= 0 Then
'    lblWCBNo.Caption = wcb(X% + 1, 1)
'    lblCase.Caption = wcb(X% + 1, 2)
'End If

End Sub


'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

'Data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'Data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove Binding Control

Call Display_Value
Data1.Refresh

'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HROHSCOS", "Cancel")
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

End Sub

'Private Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdContact_Click()
'frmEHSContact.Show
'Unload Me
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x%

On Error GoTo Del_Err
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "No Records Found"
    Exit Sub
End If

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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HROHSCOS", "Delete")
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

'Private Sub cmdIncident_Click()
'frmEHSINCIDENT.Show
'Unload Me
'End Sub

'Private Sub cmdInjLoc_Click()
'frmEHSINJURY.Show
'Unload Me
'End Sub

Sub cmdModify_Click()
Dim x%

'If Not gSec_Upd_Health_Safety Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If

On Error GoTo Mod_Err

Call SET_UP_MODE
'Call ST_UPD_MODE(True)
dlpReoccur.SetFocus

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HROHSCOS", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

'Private Sub cmdModify_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String

'If Not gSec_Upd_Health_Safety Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err
'If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    Me.vbxTrueGrid.Enabled = False
'    Data1.RecordSource = "HROHSCOS"
'    Data1.Refresh
'    fglbEmptyNew = True
'End If

Me.vbxTrueGrid.Enabled = False

'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control
Call Set_Control("B", Me)


If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

wcb(1, 1) = ""

lblCNum.Caption = "001"
medShifts = 0
'cmbWCB.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HROHSCOS", "Add")
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
Dim x%

On Error GoTo Add_Err

If Not chkHSWCBCs() Then Exit Sub
rsDATA.Requery
If fglbNew Then rsDATA.AddNew
Call UpdUStats(Me) ' update user's stats (who did it and when)

'rsDATA!CC_FUTECOA = medFutureE
'
'x% = cmbWCB.ListIndex
'If x% >= 0 Then
'    lblWCBNo.Caption = wcb(x% + 1, 1)
'    lblCase.Caption = wcb(x% + 1, 2)
'End If


If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    Call Set_Control("U", Me, rsDATA)
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    Call Set_Control("U", Me, rsDATA)
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
fglbNew = False
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

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_REOCCURENCE", "Update")
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
Dim RHeading As String

RHeading = lblEEName & "'s WSIB Cost"
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

RHeading = lblEEName & "'s WSIB Cost"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub



'Private Sub cmdTCause_Click()
'frmEHSCause.Show
'Unload Me
'End Sub

'Private Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub


Function EERetrieve()
Dim SQLQ As String
EERetrieve = False
On Error GoTo EERError
If glbtermopen Then
    SQLQ = "SELECT * from Term_OHS_REOCCURENCE "
    SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq & " ORDER BY CC_REOCCURDATE,CC_WCBNBR"
Else
    SQLQ = "SELECT * from HR_OHS_REOCCURENCE "
    SQLQ = SQLQ & "WHERE CC_EMPNBR = " & glbLEE_ID & " ORDER BY CC_REOCCURDATE,CC_WCBNBR"
End If
Data1.RecordSource = SQLQ
Data1.Refresh

Call popComboWCB

EERetrieve = True
Exit Function
EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HROHSCOS", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
Exit Function
End Function
Private Sub Form_Activate()
glbOnTop = "frmEHSReOccur"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "frmEHSReOccur"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

ReDim wcb(1, 3) 'laura nov 14, 1997
glbOnTop = "frmEHSReOccur"
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

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

If Len(glbLEE_SName) < 1 Then Exit Sub
Call popComboWCB

Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "WSIB Cost - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call ST_UPD_MODE(True)
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
Set frmEHSReOccur = Nothing ' carmen may 00
Call NextForm
End Sub


Private Sub medFutureE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medFutureES_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medHealth_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medNonE_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medNonE_ValidationError(InvalidText As String, StartPosition As Integer)
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medOther_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medPartial2_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medPartial4_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medPension_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medReEmp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub



Private Sub medRehab_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medRetPen_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medSurvivor_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub


Private Sub medTemp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub popComboWCB()
Dim snapWCBs As New ADODB.Recordset
Dim x%, SQLQ As String
    SQLQ = "SELECT DISTINCT EC_WCBNBR " 'FROM HR_OCC_HEALTH_SAFETY WHERE EC_EMPNBR = " & glbLEE_ID & " "
    If glbtermopen Then
        SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        SQLQ = SQLQ & " and EC_WCBNBR <> ' ' "
        SQLQ = SQLQ & "ORDER BY EC_WCBNBR "
        snapWCBs.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Else
        SQLQ = SQLQ & "FROM HR_OCC_HEALTH_SAFETY WHERE EC_EMPNBR = " & glbLEE_ID & " "
        SQLQ = SQLQ & " and EC_WCBNBR <> ' ' "
        SQLQ = SQLQ & "ORDER BY EC_WCBNBR "
        If snapWCBs.State <> 0 Then snapWCBs.Close
        snapWCBs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
    
    cmbWCB.Clear

    Do While Not snapWCBs.EOF
        If Not IsNull(snapWCBs("EC_WCBNBR")) Then
        cmbWCB.AddItem snapWCBs("EC_WCBNBR")
        End If
        snapWCBs.MoveNext
    Loop
    snapWCBs.Close
    
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
If Not fglbNew Then
    cmbWCB.Enabled = False
Else
    cmbWCB.Enabled = TF
End If
dlpReoccur.Enabled = TF
dlpApprDate.Enabled = TF
dlpReturnDate.Enabled = TF
medShifts.Enabled = TF
memComments.Enabled = TF

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'   cmdModify.Enabled = False
End If

End Sub

Private Sub lblClaimData_Change()
If cmbWCB.ListCount > 0 Then
    If Len(lblClaimData) = 0 Then
        cmbWCB.ListIndex = -1
    Else
        cmbWCB.Text = lblClaimData
    End If
End If
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
        
        If glbtermopen Then
            SQLQ = "SELECT * from Term_OHS_REOCCURENCE "
            SQLQ = SQLQ & "WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "SELECT * from HR_OHS_REOCCURENCE "
            SQLQ = SQLQ & "WHERE CC_EMPNBR = " & glbLEE_ID
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
 '   If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
 '   End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String, x%, WCBN$, WCBN2$

On Error GoTo Tab1_Err
Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    'MsgBox "No Records Found"
End If
Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HROHSCOS", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
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
        Call SET_UP_MODE
        'Me.cmdModify_Click
        Exit Sub
    End If

    
If glbtermopen Then
    SQLQ = "SELECT * from Term_OHS_REOCCURENCE "
    SQLQ = SQLQ & "WHERE CC_WCBC_ID = " & Data1.Recordset!CC_WCBC_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "SELECT * from HR_OHS_REOCCURENCE "
    SQLQ = SQLQ & "WHERE CC_WCBC_ID = " & Data1.Recordset!CC_WCBC_ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE
'Me.cmdModify_Click
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
Me.vbxTrueGrid.Enabled = True
Call ST_UPD_MODE(TF)
End Sub
Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEHSReOccur.Caption = " Reoccurrence - " & Left$(glbLEE_SName, 5)
    frmEHSReOccur.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
End Sub
Function chkNumeric(Ctrol As Control)
chkNumeric = False
If Len(Ctrol) = 0 Then
    Ctrol = 0
Else
    If Not IsNumeric(Ctrol.Text) Then
        MsgBox Mid(Ctrol.Tag, 4) & " Must be numeric"
        Ctrol.SetFocus
        Exit Function
    End If
End If
chkNumeric = True
End Function
