VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSRequisition 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Requisition Security"
   ClientHeight    =   9450
   ClientLeft      =   465
   ClientTop       =   1410
   ClientWidth     =   11220
   ControlBox      =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   9450
   ScaleWidth      =   11220
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LDATE"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   7740
      MaxLength       =   25
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LTIME"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   8880
      MaxLength       =   25
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8700
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LUSER"
      DataSource      =   "Data2"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   9960
      MaxLength       =   25
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   8730
      Visible         =   0   'False
      Width           =   900
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11220
      _Version        =   65536
      _ExtentX        =   19791
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
      Begin VB.Label lblPosl 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   135
         Width           =   660
      End
      Begin VB.Label lblUSERID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "USERID"
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
         TabIndex        =   19
         Top             =   120
         Width           =   870
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "USERNAME"
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
         Left            =   3030
         TabIndex        =   18
         Top             =   120
         Width           =   1290
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   23
      Top             =   8790
      Width           =   11220
      _Version        =   65536
      _ExtentX        =   19791
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
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   870
         TabIndex        =   11
         Tag             =   "Edit the information "
         Top             =   180
         Width           =   765
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   30
         TabIndex        =   12
         Tag             =   "Close and exit this screen"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1725
         TabIndex        =   7
         Tag             =   "Save the changes made"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2565
         TabIndex        =   8
         Tag             =   "Cancel the changes made"
         Top             =   180
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3450
         TabIndex        =   9
         Tag             =   "Add a new Record"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4305
         TabIndex        =   10
         Tag             =   "Delete the Record Selected"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5145
         TabIndex        =   13
         Tag             =   "Print Listing "
         Top             =   180
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6240
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
      Begin MSAdodcLib.Adodc Data2 
         Height          =   330
         Left            =   7560
         Top             =   30
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   7560
         Top             =   390
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
         LockType        =   1
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fsRequisition.frx":0000
      Height          =   3285
      Left            =   240
      OleObjectBlob   =   "fsRequisition.frx":0014
      TabIndex        =   0
      Top             =   510
      Width           =   10770
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "RS_STATUS"
      Height          =   285
      Index           =   0
      Left            =   2640
      TabIndex        =   4
      Tag             =   "01-Position Status - Code "
      Top             =   5070
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBST"
      MaxLength       =   6
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "RS_GRPCD"
      Height          =   285
      Index           =   2
      Left            =   2640
      TabIndex        =   3
      Tag             =   "01-Position Group Code "
      Top             =   4725
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "RS_POSTYPE"
      Height          =   285
      Index           =   3
      Left            =   2640
      TabIndex        =   1
      Tag             =   "00-Position Type Code"
      Top             =   4020
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "POTY"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "RS_ORG"
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Tag             =   "01-Union - Code"
      Top             =   4365
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpJobExcl 
      DataField       =   "RS_EXCLJOB"
      Height          =   285
      Left            =   2640
      TabIndex        =   5
      Tag             =   "01-Position code"
      Top             =   5760
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpJobIncl 
      DataField       =   "RS_INCLJOB"
      Height          =   285
      Left            =   2640
      TabIndex        =   6
      Tag             =   "01-Position code"
      Top             =   6480
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
      MultiSelect     =   -1  'True
   End
   Begin VB.Label lblPosGroup 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
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
      Left            =   450
      TabIndex        =   30
      Top             =   4770
      Width           =   1035
   End
   Begin VB.Label lblhelptxt2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "List positions to include which are not otherwise part of any of the groups in the table above."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3000
      TabIndex        =   29
      Top             =   7080
      Width           =   7935
   End
   Begin VB.Label lblhelptxt1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "You must choose one of the rows in the table above."
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3000
      TabIndex        =   28
      Top             =   6840
      Width           =   4515
   End
   Begin VB.Label lblJobCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Exclude Position Codes"
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
      Left            =   450
      TabIndex        =   27
      Top             =   5805
      Width           =   1665
   End
   Begin VB.Label lblJobCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Include Job Codes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   26
      Top             =   6525
      Width           =   1320
   End
   Begin VB.Label lblPosStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Status"
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
      Left            =   450
      TabIndex        =   25
      Top             =   5115
      Width           =   1050
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
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
      Left            =   450
      TabIndex        =   24
      Top             =   4410
      Width           =   960
   End
   Begin VB.Label lblPosType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Type"
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
      Left            =   450
      TabIndex        =   22
      Top             =   4065
      Width           =   960
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CompNo"
      DataField       =   "COMPNO"
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
      Left            =   10080
      TabIndex        =   21
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_Return 
         Caption         =   "&Return to Security"
      End
   End
End
Attribute VB_Name = "frmSRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsDATA As New ADODB.Recordset
Dim fglbNew As Boolean

Private Function chkSecRequisition()

Dim SQLQ As String, Msg As String, RSID&, PosType$

chkSecRequisition = False

On Error GoTo chkSecRequisition_Err

'If Len(clpCode(3).Text) = 0 Then
'    MsgBox lStr("Position Type Code is required")
'    clpCode(3).SetFocus
'    Exit Function
'End If

If clpCode(3).Caption = "Unassigned" Then
    MsgBox lStr("Position Type Code must be valid")
    clpCode(3).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" And clpCode(1).Text <> "-NON" And clpCode(1).Text <> "-EXE" Then
    MsgBox lblUnion.Caption & " must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(2).Caption = "Unassigned" Then
    MsgBox lblPosGroup.Caption & " must be valid"
    clpCode(2).SetFocus
    Exit Function
End If
If clpCode(0).Caption = "Unassigned" Then
    MsgBox lblPosStatus.Caption & " must be valid"
    clpCode(0).SetFocus
    Exit Function
End If


If Not clpJobExcl.ListChecker Then
    Exit Function
End If

If Not clpJobIncl.ListChecker Then
    Exit Function
End If

If Len(clpJobExcl.Text) > 500 Then
    MsgBox "Exclude Job Codes cannot exceed 500 characters"
    clpJobExcl.SetFocus
    Exit Function
End If

If Len(clpJobIncl.Text) > 500 Then
    MsgBox "Include Job Codes cannot exceed 500 characters"
    clpJobIncl.SetFocus
    Exit Function
End If

If Len(clpCode(3).Text) = 0 And Len(clpCode(2).Text) = 0 And Len(clpCode(1).Text) = 0 And Len(clpCode(0).Text) = 0 And Len(clpJobExcl.Text) = 0 And Len(clpJobIncl.Text) = 0 Then
    MsgBox "Requisition Security record cannot be blank"
    clpCode(3).SetFocus
    Exit Function
End If

If fglbNew Then
    If IsNull(rsDATA!RS_ID) Then RSID& = 0 Else RSID& = Val(rsDATA!RS_ID)
    PosType$ = clpCode(3).Text
    If modISDupRequisition(glbSecUSERID, clpCode(3).Text, clpCode(1).Text, clpCode(2).Text, clpCode(0).Text, RSID&, clpJobIncl.Text, clpJobExcl.Text) Then
        MsgBox lStr("Position Type, Union, Position Group, Position Status, and Include Job Codes, Exclude Job Codes must be unique")
        clpCode(3).SetFocus
        Exit Function
    End If
End If

chkSecRequisition = True

Exit Function

chkSecRequisition_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSecRequisition", "HRA_SECURE_REQUISITION", "validation")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub cmdCancel_Click()

On Error GoTo Can_Err

fglbNew = False

rsDATA.CancelUpdate

Call Display_Value

Call ST_UPD_MODE(False)  ' reset screen's attributes

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HRA_SECURE_REQUISITION", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

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
Dim a As Integer, Msg As String, INo&

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "


a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HRA_SECURE_REQUISITION", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdDelete_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()
Dim SQLQ As String

Call ST_UPD_MODE(True)

On Error GoTo Edit_Err

clpCode(3).SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HRA_SECURE_REQUISITION", "Modify")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdNew_Click()
Dim SQLQ As String

'Ticket #18668 - Remove the limit - esp. for Linamar
'I tested the WHERE clause of the Dept security and it worked fine and it also worked for reports
'as well - I tried to setup 71 depts security rows and it worked fine. Also searched internet -
'nowhere I could find any limit to WHERE clause.
'So we are not sure why this was added
'If Data1.Recordset.RecordCount = 50 Then
'    MsgBox "You can't add more than 50 departments"
'    Exit Sub
'End If

Call ST_UPD_MODE(True)

On Error GoTo AddN_Err

Call Set_Control("B", Me)

rsDATA.AddNew

clpCode(3).Caption = ""
lblCNum.Caption = "001"
lblUSERID.Caption = glbSecUSERID

fglbNew = True

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRA_SECURE_REQUISITION", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub CmdNew_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim x%
Dim xID
Dim xTemplate As String

On Error GoTo OK_Err

''Ticket #20585 - If Template then update users with this template as well.
''If User and with no template then update that user's profile.
''if User and with Template then do not update user's profile.
''Get the Template Name of this User ID
''xTemplate = Get_Template(glbSecUSERID)
'
'If xTemplate = "TEMPLATE" Then
'    'Update all users with this template. After the changes are saved
'ElseIf xTemplate = "" Then
'    'User - User with no template - don't do anything let system update user's profile
'ElseIf xTemplate <> "TEMPLATE" Then
'    'User with template - do not allow to save these changes.
'    MsgBox "Security change cannot be saved. This user's security profile is based on the '" & xTemplate & "' template.", vbExclamation, "Template based User Security Profile"
'End If

''if Template or User
'If xTemplate = "TEMPLATE" Or xTemplate = "" Then

    If Not chkSecRequisition() Then Exit Sub
    
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    
    Call Set_Control("U", Me, rsDATA)
    
    rsDATA("USERID") = lblUSERID & ""
    rsDATA("RS_INCLJOB") = clpJobIncl.Text
    rsDATA("RS_EXCLJOB") = clpJobExcl.Text
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
    
    gdbAdoIhr001.Execute "UPDATE HRA_SECURE_REQUISITION SET RS_INCLJOB = '" & clpJobIncl.Text & "' WHERE USERID ='" & Replace(lblUSERID, "'", "''") & "'"
'End If

Data1.Refresh

'Ticket #21629 - Jerry said not to change user's Dept security based on Template's Dept Security
'Ticket #20585 - Security Based on Template Profile
'If xTemplate = "TEMPLATE" Then
'    'Call procedure to Update all users with this template.
'    Call Update_Users_withthis_Template(glbSecUSERID)
'End If

fglbNew = False

Call ST_UPD_MODE(False)

Me.vbxTrueGrid.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRA_SECURE_REQUISITION", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName.Caption & "'s Requisition Security"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRA_SECURE_REQUISITION", "SELECT")

End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%
Dim xTemplate  As String
Dim SQLQ As String

Screen.MousePointer = HOURGLASS

lblUSERID.Caption = glbSecUSERID
lblEEName.Caption = glbSecEEName

Me.Caption = lStr("Requisition Security - ") & lblEEName

'Ticket #29024 - Had to comment this so this form can be shown Modal form from frmSecurity
'frmSRequisition.Show

Data1.ConnectionString = glbAdoIHRDB

x% = modGetRequisitions()

''????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
'xTemplate = ""
'xTemplate = Get_Template(glbSecUSERID)

'SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION "
'If xTemplate = "" Or xTemplate = "TEMPLATE" Then
'    SQLQ = SQLQ & " WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "'"
'Else
'    '????Ticket #24808 -  Retrieve template's security profile
'    SQLQ = SQLQ & " WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
'End If
'SQLQ = SQLQ & " ORDER BY RS_POSTYPE"
'
'Data1.RecordSource = SQLQ
'Data1.Refresh

Call setCaption(lblPosType)
Call setCaption(lblUnion)
Call setCaption(lblPosGroup)
Call setCaption(lblPosStatus)

Call setCaption(Me.vbxTrueGrid.Columns(0))
Call setCaption(Me.vbxTrueGrid.Columns(1))
Call setCaption(Me.vbxTrueGrid.Columns(2))
Call setCaption(Me.vbxTrueGrid.Columns(3))

If vbxTrueGrid.Visible Then
    Me.vbxTrueGrid.SetFocus
End If

Call INI_Controls(Me)

Call ST_UPD_MODE(False)

'Ticket #21629 - Jerry said Department Security is independent of Template Security
'Ticket #20585 - Enable/Disable Edit, New and Delete buttons based on the type of user
'xTemplate = Get_Template(glbSecUSERID)
'If xTemplate = "" Or xTemplate = "TEMPLATE" Then
'    'User without Template or Template
'Else
'    'User with Template
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'End If

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmSRequisition = Nothing
End Sub

Private Sub mnu_File_Exit_Click()
    Call ApplicationEnd
End Sub

Private Sub mnu_F_PrintSetup_Click()
MDIMain.vbxCommon.Action = 5
End Sub

Private Sub mnu_Return_Click()
   Unload Me
End Sub

Private Function modGetRequisitions()
Dim SQLQ

modGetRequisitions = False

Screen.MousePointer = HOURGLASS

On Error GoTo modGetRequisitionsErr

SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION"
SQLQ = SQLQ & " WHERE USERID = '" & Replace(glbSecUSERID, "'", "''") & "'"
SQLQ = SQLQ & " ORDER BY RS_POSTYPE"

Data1.RecordSource = SQLQ
Data1.Refresh


modGetRequisitions = True
Screen.MousePointer = DEFAULT

Exit Function

modGetRequisitionsErr:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Security Requisitions", "HRA_SECURE_REQUISITION", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Exit Function

End Function

Private Function modISDupRequisition(UserID As String, PosType As String, Org As String, PosGroup As String, PosStatus As String, RSID As Long, IncludeJob As String, ExcludeJob As String)
Dim SQLQ As String
Dim snapDepSec As New ADODB.Recordset

modISDupRequisition = True

On Error GoTo modISDupRequisition_Err

Screen.MousePointer = HOURGLASS

SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION "
SQLQ = SQLQ & " WHERE USERID = '" & Replace(UserID, "'", "''") & "'"
If Len(PosType) > 0 Then
    SQLQ = SQLQ & " AND RS_POSTYPE = '" & PosType & "' "
Else
    SQLQ = SQLQ & " AND RS_POSTYPE IS NULL "
End If

SQLQ = SQLQ & " AND RS_ID <> " & RSID

If Len(Org) > 0 Then
    SQLQ = SQLQ & " AND RS_ORG = '" & Org & "' "
Else
    SQLQ = SQLQ & " AND RS_ORG IS NULL "
End If
If Len(PosGroup) > 0 Then
    SQLQ = SQLQ & " AND RS_GRPCD = '" & PosGroup & "' "
Else
    SQLQ = SQLQ & " AND RS_GRPCD IS NULL "
End If
If Len(PosStatus) > 0 Then
    SQLQ = SQLQ & " AND RS_STATUS = '" & PosStatus & "' "
Else
    SQLQ = SQLQ & " AND RS_STATUS IS NULL "
End If

If Len(IncludeJob) > 0 Then
    SQLQ = SQLQ & " AND RS_INCLJOB LIKE '" & getEmpnbr(IncludeJob) & "' "
'Else
'    SQLQ = SQLQ & " AND RS_INCLJOB IS NULL "
End If
If Len(ExcludeJob) > 0 Then
    SQLQ = SQLQ & " AND RS_EXCLJOB LIKE '" & getEmpnbr(ExcludeJob) & "' "
'Else
'    SQLQ = SQLQ & " AND RS_EXCLJOB IS NULL "
End If

snapDepSec.Open SQLQ, gdbAdoIhr001, adOpenStatic
If snapDepSec.BOF And snapDepSec.EOF Then
    modISDupRequisition = False
End If

snapDepSec.Close
Screen.MousePointer = DEFAULT

Exit Function

modISDupRequisition_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Find Duplicate", "HRA_SECURE_REQUISITION", "SELECT")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

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
    cmdNew.Enabled = FT
    cmdModify.Enabled = FT
    cmdDelete.Enabled = FT
    cmdPrint.Enabled = FT
    clpCode(3).Enabled = TF     'Position  Type
    clpCode(1).Enabled = TF     'Union
    clpCode(2).Enabled = TF     'Position Group
    clpCode(0).Enabled = TF     'Position Status
    vbxTrueGrid.Enabled = FT
    clpJobIncl.Enabled = TF
    clpJobExcl.Enabled = TF
        
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        cmdModify.Enabled = False
    End If
    
    If Not gSec_Upd_Security Then 'And Not gSec_Upd_Quick_ESS Then
        cmdModify.Enabled = False
        cmdNew.Enabled = False
        cmdDelete.Enabled = False
    End If
    
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
    Dim SQLQ As String
    Dim xTemplate As String

    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
    xTemplate = ""
    xTemplate = Get_Template(glbSecUSERID)

    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION "
    
    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
        SQLQ = SQLQ & " WHERE USERID='" & Replace(glbSecUSERID, "'", "''") & "'"
    Else
        '????Ticket #24808 -  Retrieve template's security profile
        SQLQ = SQLQ & " WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
    End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    
    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Then ' if the tab key was struck
        KeyAscii = 0
        If cmdOK.Enabled Then
            cmdOK.SetFocus
        Else
            cmdClose.SetFocus
        End If
    End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, x%

Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value
Exit Sub

Tab1_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRA_SECURE_REQUISITION", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
    
    SQLQ = "SELECT * from HRA_SECURE_REQUISITION "
    SQLQ = SQLQ & " WHERE RS_ID = " & Data1.Recordset!RS_ID

    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
         
End Sub

Private Sub Update_Users_withthis_Template(xTemplate)
    Dim SQLQ As String
    Dim rsSecBasic As New ADODB.Recordset
    
    'Retrieve all users associated with this changed Template
    SQLQ = "SELECT USERID, SECURE_TEMPLATE FROM HR_SECURE_BASIC WHERE SECURE_TEMPLATE = '" & xTemplate & "'"
    SQLQ = SQLQ & " ORDER BY USERID"
    rsSecBasic.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSecBasic.EOF
        If Not IsNull(rsSecBasic("USERID")) Then
            'Update each user with this changed Template
            Call SpecificFunction_Template_Based_Security_Profile_Update(rsSecBasic("USERID"), xTemplate, "Change", "REQUISITION")
        End If
        rsSecBasic.MoveNext
    Loop
    rsSecBasic.Close
    Set rsSecBasic = Nothing
    
End Sub

