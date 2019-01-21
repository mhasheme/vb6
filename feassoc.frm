VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEASSOC 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Associations"
   ClientHeight    =   8160
   ClientLeft      =   -150
   ClientTop       =   765
   ClientWidth     =   9540
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
   ScaleHeight     =   8160
   ScaleWidth      =   9540
   WindowState     =   2  'Maximized
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "TD_COMMENTS"
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
      Left            =   330
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "00-Associations Comments"
      Top             =   5220
      Width           =   8805
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   8340
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   855
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feassoc.frx":0000
      Height          =   2085
      Left            =   90
      OleObjectBlob   =   "feassoc.frx":0014
      TabIndex        =   0
      Tag             =   "Listing of Associations"
      Top             =   750
      Width           =   9075
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "TD_RENEWDT"
      Height          =   315
      Index           =   1
      Left            =   1620
      TabIndex        =   5
      Tag             =   "40-Membership's renewal date"
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "TD_BEGINDT"
      Height          =   315
      Index           =   0
      Left            =   1620
      TabIndex        =   4
      Tag             =   "41-Membership's effective starting date"
      Top             =   4080
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   556
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "TD_CODE"
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   1
      Tag             =   "01-Association/Membership- Code"
      Top             =   3000
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "TDCD"
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6240
      Top             =   7560
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
      TabIndex        =   21
      Top             =   7500
      Width           =   9540
      _Version        =   65536
      _ExtentX        =   16828
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
         Left            =   9810
         Top             =   210
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
   Begin VB.CheckBox chkCompPaid 
      Alignment       =   1  'Right Justify
      Caption         =   "Company Paid       "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   300
      TabIndex        =   3
      Tag             =   "Company Paid "
      Top             =   3660
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.TextBox txtCompPaid 
      Appearance      =   0  'Flat
      DataField       =   "TD_COMPPD"
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
      Left            =   2640
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TD_LDATE"
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
      TabIndex        =   7
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TD_LTIME"
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
      Left            =   4680
      MaxLength       =   25
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "TD_LUSER"
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
      Left            =   6360
      MaxLength       =   25
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9540
      _Version        =   65536
      _ExtentX        =   16828
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
         TabIndex        =   22
         Top             =   135
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
         TabIndex        =   14
         Top             =   155
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
         Left            =   1440
         TabIndex        =   13
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
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
         Left            =   3100
         TabIndex        =   12
         Top             =   135
         Width           =   1740
      End
   End
   Begin MSMask.MaskEdBox medDuesPaid 
      DataField       =   "TD_DUES"
      Height          =   285
      Left            =   1940
      TabIndex        =   2
      Tag             =   "20-Enter $ amount of dues paid"
      Top             =   3330
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
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
      Format          =   "$##,##0.00;($##,##0.00)"
      PromptChar      =   "_"
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
      Index           =   5
      Left            =   330
      TabIndex        =   25
      Top             =   4920
      Width           =   2910
   End
   Begin VB.Image imgNoSec 
      Height          =   240
      Left            =   7920
      Picture         =   "feassoc.frx":4234
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSec 
      Height          =   240
      Left            =   7920
      Picture         =   "feassoc.frx":437E
      Top             =   3000
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblImport 
      Alignment       =   1  'Right Justify
      Caption         =   "Associations"
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
      Left            =   6000
      TabIndex        =   24
      Top             =   3000
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Renewal Date"
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
      Index           =   4
      Left            =   330
      TabIndex        =   20
      Top             =   4500
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Starting Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   19
      Top             =   4140
      Width           =   1140
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dues Paid"
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
      Left            =   330
      TabIndex        =   18
      Top             =   3375
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Associations"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   17
      Top             =   3045
      Width           =   1350
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "TD_EMPNBR"
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
      Left            =   1920
      TabIndex        =   15
      Top             =   7200
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "TD_COMPNO"
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
      Left            =   240
      TabIndex        =   16
      Top             =   7200
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEASSOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim fglHredsem As String
Dim oldCode As String
Dim oldBeginDt As Date

Private Sub chkCompPaid_Click()
If chkCompPaid.Value = 1 Then
    txtCompPaid.Text = "Y"
Else
    txtCompPaid.Text = "N"
End If
End Sub

Private Sub chkCompPaid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function chkEASSOC()
Dim oCode As String, OCodeD As String

chkEASSOC = False

On Error GoTo chkEASSOC_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Association code is a required field"
    clpCode(1).SetFocus
    Exit Function
End If

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Association code must be valid"
    clpCode(1).SetFocus
    Exit Function
End If

If chkCompPaid.Value = 1 Then
    txtCompPaid.Text = "Y"
Else
    txtCompPaid.Text = "N"
End If

If Len(dlpDate(0).Text) < 1 Then
    MsgBox "Starting Date is Required Field"
    dlpDate(0).SetFocus
    Exit Function
Else
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Starting Date is not a valid date."
        dlpDate(0).SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(1).Text) > 0 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Renewal Date is not a valid date."
        dlpDate(1).Text = ""
        dlpDate(1).SetFocus
        Exit Function
    End If
End If

If Len(Trim(medDuesPaid)) = 0 Then
    medDuesPaid = 0
Else
    If Not IsNumeric(medDuesPaid) Then
        MsgBox "Dues Paid must be numeric"
        medDuesPaid.SetFocus
        Exit Function
    End If
End If

If Len(memComments) > 4000 Then
    MsgBox "The Comments field can only contain 4000 charaters. It is exceeding 4000 characters."
    memComments.SetFocus
    Exit Function
End If

chkEASSOC = True

Exit Function

chkEASSOC_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkAssoc", "HRTRADE", "edit/Add")
Call RollBack '23July99 js

End Function

Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Display_Value


fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(True)  ' reset screen's attributes

fglbNew = False
Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRTRADE", "Cancel")
Call RollBack '23July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEASSOC" Then glbOnTop = ""

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

fglHredsem = dlpDate(1).Text
If fglHredsem <> "" Then
    If Not updFollow("D") Then
        Exit Sub
    End If
End If

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " AND TD_CODE='" & clpCode(1).Text & "' AND TD_BEGINDT=" & Date_SQL(dlpDate(0).Text)  '" and TD_DOCKEY=" & glbDocKey & " " '
    End If
  
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR = " & glbLEE_ID & " AND TD_CODE='" & clpCode(1).Text & "' AND TD_BEGINDT=" & Date_SQL(dlpDate(0).Text)   '" and TD_DOCKEY=" & glbDocKey & " "
    End If
    
    Data1.Refresh
End If

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRTRADE", "Delete")
Call RollBack '23July99 js

End Sub

Sub cmdNew_Click()
Dim SQLQ As String

On Error GoTo AddN_Err

fglbNew = True
Call SET_UP_MODE

If gsAttachment_DB Then
    lblImport.Visible = True 'False
    imgSec.Visible = False
    imgNoSec.Visible = True 'False
    cmdImport.Visible = True 'False
End If

Call Set_Control("B", Me)

rsDATA.AddNew

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

lblCNum.Caption = "001"
clpCode(1).Enabled = True
clpCode(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err


Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRTRADE", "Add")
Call RollBack '23July99 js

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
    Dim X, xID
    On Error GoTo Add_Err
    
    If Not chkEASSOC() Then Exit Sub
    
    Call UpdUStats(Me) ' update user's stats (who did it and when)
    
    Call Set_Control("U", Me, rsDATA)
    
    If glbtermopen Then
        rsDATA!TERM_SEQ = glbTERM_Seq
        gdbAdoIhr001X.BeginTrans
        rsDATA.Update
        gdbAdoIhr001X.CommitTrans
        
        If gsAttachment_DB Then
            gdbAdoIhr001_DOC.Execute "Update Term_HRDOC_TRADE set TD_CODE='" & rsDATA("TD_CODE") & "',TD_BEGINDT=" & Date_SQL(rsDATA("TD_BEGINDT")) & " WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " AND TD_CODE='" & oldCode & "' AND TD_BEGINDT=" & Date_SQL(oldBeginDt) & " AND TD_DOCKEY = " & glbDocKey
        End If
    Else
        gdbAdoIhr001.BeginTrans
        rsDATA.Update
        gdbAdoIhr001.CommitTrans
        
        If gsAttachment_DB Then
            If Not fglbNew Then
                gdbAdoIhr001_DOC.Execute "Update HRDOC_TRADE set TD_CODE='" & rsDATA("TD_CODE") & "',TD_BEGINDT=" & Date_SQL(rsDATA("TD_BEGINDT")) & " WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR = " & glbLEE_ID & " AND TD_CODE='" & oldCode & "' AND TD_BEGINDT=" & Date_SQL(oldBeginDt) & " AND TD_DOCKEY = " & glbDocKey
            End If
        End If
    End If
    xID = rsDATA!TD_ID
    Data1.Refresh
    Data1.Recordset.Find "TD_ID=" & xID
    
    'Ticket #22682: Release 8.0 - Set older Performance Review Follow Up records as Completed first if uncompleted
    'follow up records are found for Salary, before adding a new follow up record.
    If fglbNew Then
        glbFollowUpList = "AREN"
        If Older_FollowUp_Records_Found(glbFollowUpList) Then
            frmFollowUpList.Show 1
        End If
    End If
    
    If Not updFollow("U") Then Exit Sub
    
    'Call ST_UPD_MODE(True)
    
    fglbNew = False
    
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
    
    Call SET_UP_MODE
    
    Me.vbxTrueGrid.SetFocus
    
    If NextFormIF("Association") Then
        Call cmdNew_Click
    End If
    
    Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRTRADE", "Update")
Call RollBack '23July99 js

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s " & lStr("Associations") & "/Memberships"
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

RHeading = lblEEName & "'s " & lStr("Associations") & "/Memberships"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub
'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Function EERetrieve()
Dim SQLQ As String

Screen.MousePointer = HOURGLASS

EERetrieve = False

On Error GoTo EERError

If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_TRADE"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY TD_CODE"
Else
    SQLQ = "Select * from HRTRADE"
    SQLQ = SQLQ & " where TD_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY TD_CODE"
End If

Data1.RecordSource = SQLQ
Data1.Refresh
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SklsRetrieve", "HRTrade", "SELECT")
Call RollBack '23July99 js

Exit Function
End Function

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "Associations"
    If fglbNew Then
        glbDocKey = 0
        If Len(dlpDate(0).Text) = 0 Or Len(clpCode(1).Text) = 0 Then
            MsgBox "'Associations' and 'Starting Date' must be entered before attaching a document", vbExclamation
            Exit Sub
        Else
            glbAssocCode = clpCode(1).Text
            glbBeginDt = dlpDate(0).Text
        End If
    Else
        glbDocKey = rsDATA("TD_ID") 'Ticket #16018
        glbAssocCode = rsDATA("TD_CODE")
        glbBeginDt = rsDATA("TD_BEGINDT")
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmEASSOC")
End Sub

Private Sub Form_Activate()
    glbOnTop = "FRMEASSOC"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEASSOC"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMEASSOC"
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
    Me.Caption = "Associations/Memberships - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If

lblEENum.Caption = ShowEmpnbr(lblEEID)

Call Display_Value
Call ST_UPD_MODE(True)             '
'If Not gSec_Upd_Associations Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'End If
Call INI_Controls(Me)

lblTitle(1).Caption = lStr(lblTitle(1).Caption)
vbxTrueGrid.Columns(0).Caption = lStr(vbxTrueGrid.Columns(0).Caption)
lblImport.Caption = lStr(lblImport.Caption)

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
    Call NextForm
End Sub

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmEASSOC")
    Call FillMemoFile(SQLQ, "Associations")
End Sub

Private Sub medDuesPaid_GotFocus()
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

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT

chkCompPaid.Enabled = TF
medDuesPaid.Enabled = TF
clpCode(1).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
memComments.Enabled = TF
memComments.Locked = FT

'''chkCompPaid.Enabled = FT
'''medDuesPaid.Enabled = FT
'''clpCode(1).Enabled = FT
'''dlpDate(0).Enabled = FT
'''dlpDate(1).Enabled = FT

'cmdCancel.Enabled = TF
'cmdOK.Enabled = TF
'vbxTrueGrid.Enabled = FT

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then

End If

glbDocName = "Associations"
If gsAttachment_DB Then
    If Not (rsDATA.BOF And rsDATA.EOF) Then
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("TD_DOCKEY")) Then
                glbDocKey = rsDATA("TD_DOCKEY")
            Else
                glbDocKey = 0
            End If
        Else
            If Not IsNull(Data1.Recordset("TD_DOCKEY")) Then
                glbDocKey = Data1.Recordset("TD_DOCKEY")
            Else
                glbDocKey = 0
            End If
        End If
    
        glbAssocCode = clpCode(1).Text
        If IsDate(dlpDate(0).Text) Then
            glbBeginDt = dlpDate(0).Text
        End If
    End If
    
    Call DispimgIcon(Me, "frmEASSOC")
    If gSec_Upd_Associations Then
        If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If

End Sub

Private Sub medDuesPaid_LostFocus()
If Len(Trim(medDuesPaid)) = 0 Then medDuesPaid = 0
End Sub

Private Sub memComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtCompPaid_Change()
    If txtCompPaid = "Y" Then
        chkCompPaid = 1
    Else
        chkCompPaid = 0
    End If
End Sub
'Private Sub txtDate_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtDate_DblClick(Index As Integer)
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtDate_GotFocus(Index As Integer)
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
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
        
        If glbtermopen Then         'Lucy July 5, 2000
            SQLQ = "Select * from Term_TRADE"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HRTRADE"
            SQLQ = SQLQ & " where TD_EMPNBR = " & glbLEE_ID
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
    'If cmdOK.Enabled Then
    '    cmdOK.SetFocus
    'Else
    '    cmdModify.SetFocus
    'End If
End If

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
'If Not Fnd_Match_Data2() Then Exit Sub 'MsgBox "No Records Found"

' ' set description for code
'If Data1.Recordset.RecordCount <> 0 Then
'    If Not IsNull(Data2.Recordset("TD_RENEWDT")) Then
'        txtDate(1) = Data2.Recordset("TD_RENEWDT")
'    Else
'        txtDate(1) = ""
'    End If
'End If
Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HRTrade", "Add")
Call RollBack '23July99 js

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
    Exit Sub
End If
      
If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    
    If glbtermopen Then
    SQLQ = "Select * from Term_TRADE"
    SQLQ = SQLQ & " WHERE TD_ID = " & Data1.Recordset!TD_ID
    SQLQ = SQLQ & " AND TERM_SEQ = " & Data1.Recordset!TERM_SEQ
    SQLQ = SQLQ & " ORDER BY TD_CODE"
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HRTRADE"
    SQLQ = SQLQ & " where TD_ID = " & Data1.Recordset!TD_ID
    SQLQ = SQLQ & " ORDER BY TD_CODE"
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If

If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE

fglHredsem = dlpDate(1).Text
oldCode = clpCode(1).Text
If IsDate(dlpDate(0).Text) Then
    oldBeginDt = dlpDate(0).Text
End If

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
UpdateRight = gSec_Upd_Associations
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
    frmEASSOC.Caption = "Associations - " & Left$(glbLEE_SName, 5)
    frmEASSOC.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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


Private Function updFollow(xType)   'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim rsTT As New ADODB.Recordset
Dim Edit1 As Integer

'Don't need a message for follow up - Jerry asked for v7.6

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

If fglHredsem <> "" Then    'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'AREN'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(fglHredsem)
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" Then
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If fglbNew And dlpDate(1).Text <> "" Then
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'AREN'"
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpDate(1).Text)
        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
            'Create the Code if not already existing
            rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='AREN'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                rsTT("TB_KEY") = "AREN"
                rsTT("TB_DESC") = lStr("Associations") & " Renewal"
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            Set rsTT = Nothing

            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            Call Grant_FollowUpCode_Security(glbUserID, "AREN", lStr("Associations") & " Renewal")

            'Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            rsTB("EF_FREAS") = "AREN"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And dlpDate(1).Text <> "" Then
        ' 5/2/2001 Add by Frank for no duplicated record of HR_FOLLOW_UP Begin
        SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = 'AREN' "
        SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(dlpDate(1).Text)
        

        rsFollow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsFollow.EOF Then
            'Create the Code if not already existing
            rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='AREN'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
            If rsTT.EOF Then
                rsTT.AddNew
                rsTT("TB_COMPNO") = "001"
                rsTT("TB_NAME") = "FURE"
                rsTT("TB_KEY") = "AREN"
                rsTT("TB_DESC") = lStr("Associations") & " Renewal"
                rsTT("TB_LUSER") = glbUserID
                rsTT("TB_LDATE") = Date
                rsTT("TB_LTIME") = Time$
                rsTT.Update
            End If
            rsTT.Close
            Set rsTT = Nothing
        
            'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
            'follow up record
            Call Grant_FollowUpCode_Security(glbUserID, "AREN", lStr("Associations") & " Renewal")
        
            'Add by Frank for no duplicated record of HR_FOLLOW_UP End
            rsTB.AddNew
            rsTB("EF_COMPNO") = "001"
            rsTB("EF_EMPNBR") = glbLEE_ID
            rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
            rsTB("EF_FREAS_TABL") = "FURE"
            'Ticket #24257 - Do not update Admin By for them only
            If glbCompSerial <> "S/N - 2262W" Then
                rsTB("EF_ADMINBY_TABL") = "EDAB"
                rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
            End If
            rsTB("EF_FREAS") = "AREN"
            rsTB("EF_COMMENTS") = ""
            rsTB("EF_LDATE") = Date
            rsTB("EF_LTIME") = Time$
            rsTB("EF_LUSER") = glbUserID
            rsTB.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            'Msg = "A Follow Up Record was created!"
            'MsgBox Msg
        End If
        rsFollow.Close
        rsTB.Close
        updFollow = True
        Exit Function
    End If
  
    If fglbNew = False And Edit1 = True And dlpDate(1).Text <> "" Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = dlpDate(1).Text
            dynHRAT("EF_FREAS") = "AREN"
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            ' dkostka - 02/04/2002 - Added pause to help St. Thomas db corruption problems (or try to at least)
            Call Pause(0.5)
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If fglHredsem <> dlpDate(1).Text Then
            'Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpDate(1).Text = "" Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        'Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    End If
Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
       ' Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If dlpDate(1).Text = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Associations", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function
