VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmAttendanceGroup 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Benefit Group Matrix"
   ClientHeight    =   7650
   ClientLeft      =   1125
   ClientTop       =   795
   ClientWidth     =   8985
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
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7650
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataField       =   "BM_COMMENTS"
      DataSource      =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2110
      MaxLength       =   4000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Tag             =   "00-Comments - free form"
      Top             =   5880
      Width           =   6735
   End
   Begin VB.ComboBox cmbWDate 
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
      Left            =   2110
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtCertiNum 
      Appearance      =   0  'Flat
      DataField       =   "BM_CERTIFICATE_PREFIX"
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
      Left            =   2110
      MaxLength       =   5
      TabIndex        =   5
      Tag             =   "00-Certificate Number Prefix"
      Top             =   4440
      Width           =   945
   End
   Begin VB.TextBox txtCovClass 
      Appearance      =   0  'Flat
      DataField       =   "BM_BENEFIT_CLASS"
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
      Left            =   2110
      MaxLength       =   10
      TabIndex        =   4
      Tag             =   "00-Coverage Class"
      Top             =   4080
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6600
      Top             =   7560
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
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
      TabIndex        =   15
      Top             =   6990
      Width           =   8985
      _Version        =   65536
      _ExtentX        =   15849
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5160
         TabIndex        =   36
         Tag             =   "Print Departmental Listing"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdBenAll 
         Appearance      =   0  'Flat
         Caption         =   "&Update All"
         Height          =   375
         Left            =   6960
         TabIndex        =   28
         Tag             =   "Print Division Listing"
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   135
         TabIndex        =   16
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   960
         TabIndex        =   17
         Tag             =   "Edit the information "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Tag             =   "Save changes made"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   19
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3540
         TabIndex        =   20
         Tag             =   "Create a new Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4350
         TabIndex        =   21
         Tag             =   "Delete Division listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdBenDiv 
         Appearance      =   0  'Flat
         Caption         =   "&Update Division"
         Height          =   375
         Left            =   8640
         TabIndex        =   22
         Tag             =   "Print Division Listing"
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   1935
         Top             =   30
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.TextBox txtBenAccount 
      Appearance      =   0  'Flat
      DataField       =   "BM_BENEFIT_ACCOUNT"
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
      Left            =   2110
      MaxLength       =   3
      TabIndex        =   3
      Tag             =   "00-Benefit Account"
      Top             =   3720
      Width           =   945
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BM_LDATE"
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
      Left            =   5160
      MaxLength       =   25
      TabIndex        =   12
      Text            =   "Ldate"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BM_LTIME"
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
      Left            =   6840
      MaxLength       =   25
      TabIndex        =   13
      Text            =   "LTime"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "BM_LUSER"
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
      Left            =   8520
      MaxLength       =   25
      TabIndex        =   14
      Text            =   "LUser"
      Top             =   7200
      Visible         =   0   'False
      Width           =   1590
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmAttendanceGroup.frx":0000
      Height          =   2835
      Left            =   120
      OleObjectBlob   =   "frmAttendanceGroup.frx":0014
      TabIndex        =   0
      Tag             =   "Division Listings"
      Top             =   0
      Width           =   8745
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "BM_BENEFIT_GROUP"
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Tag             =   "01-Benefit - Group Code"
      Top             =   3000
      Width           =   4140
      _ExtentX        =   7303
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "BGMF"
      MaxLength       =   10
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "BM_DIV"
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Tag             =   "01-Division - Code"
      Top             =   3360
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin VB.TextBox txtWDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      DataField       =   "BM_WHICH_DATE"
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
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   30
      Tag             =   "00-Certificate Number Prefix"
      Top             =   4800
      Visible         =   0   'False
      Width           =   465
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "BM_TO_DATE"
      DataSource      =   " "
      Height          =   285
      Index           =   1
      Left            =   6480
      TabIndex        =   8
      Tag             =   "40-Status To Date"
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "BM_FROM_DATE"
      DataSource      =   " "
      Height          =   285
      Index           =   0
      Left            =   4440
      TabIndex        =   7
      Tag             =   "40-Status From Date"
      Top             =   4800
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      ShowDescription =   0   'False
      TextBoxWidth    =   1180
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "BM_EMP"
      Height          =   285
      HelpContextID   =   2
      Index           =   2
      Left            =   1800
      TabIndex        =   9
      Tag             =   "00-Enter Status Code"
      Top             =   5160
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "BM_TERM_REASON"
      DataSource      =   " "
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   10
      Tag             =   "Termination Code - Code "
      Top             =   5520
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "TERM"
   End
   Begin VB.Label lblTitle 
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
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   35
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lblTermReason 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Termination Reason"
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
      Index           =   17
      Left            =   120
      TabIndex        =   34
      Top             =   5520
      Width           =   1425
   End
   Begin VB.Label lblEEStatus 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
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
      Left            =   120
      TabIndex        =   33
      Top             =   5160
      Width           =   1350
   End
   Begin VB.Label lblTDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Left            =   6120
      TabIndex        =   32
      Top             =   4800
      Width           =   195
   End
   Begin VB.Label lblFDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
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
      Left            =   3960
      TabIndex        =   31
      Top             =   4800
      Width           =   465
   End
   Begin VB.Label lblWDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Which Date"
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
      Left            =   120
      TabIndex        =   29
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lblCertiNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Certificate Number Prefix"
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
      Left            =   120
      TabIndex        =   27
      Top             =   4440
      Width           =   1860
   End
   Begin VB.Label lblCovclass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Coverage Class"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label lblBenAccount 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Account"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   1380
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   3360
      Width           =   690
   End
   Begin VB.Label lblBenGrp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Benefit Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   3000
      Width           =   1185
   End
End
Attribute VB_Name = "frmAttendanceGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec% ' new record
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO
Dim xTotRecCount As Long


Private Function chkBenGrp()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset
Dim x
chkBenGrp = False
On Error GoTo chkBenGrp_Err

If Len(clpDiv.Text) < 1 Then
    MsgBox lStr("Division is a required field")
    clpDiv.SetFocus
    Exit Function
End If

If Len(clpCode(1).Text) < 1 Then
    MsgBox ("Benefit Group is a required field")
    clpCode(1).SetFocus
    Exit Function
End If

If Len(txtBenAccount.Text) < 1 Then
    MsgBox ("Benefit Account is a required field")
    txtBenAccount.SetFocus
    Exit Function
End If

If Len(txtCovClass.Text) < 1 Then
    MsgBox ("Coverage Class is a required field")
    txtCovClass.SetFocus
    Exit Function
End If


If fglbNewRec% Then
    SQLQ = "SELECT * from HR_BENEFITS_GROUP_MATRIX "
    SQLQ = SQLQ & "WHERE BM_DIV = '" & clpDiv.Text & "'"
    SQLQ = SQLQ & "AND BM_BENEFIT_GROUP = '" & clpCode(1).Text & "'"
    SQLQ = SQLQ & "AND BM_BENEFIT_ACCOUNT = '" & txtBenAccount.Text & "'"
    SQLQ = SQLQ & "AND BM_BENEFIT_CLASS = '" & txtCovClass.Text & "'"
    
    If snapDivs.State <> 0 Then snapDivs.Close
    snapDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapDivs.BOF And snapDivs.EOF Then
        snapDivs.Close
    Else
        Msg$ = lStr("Duplicate record found!")
        MsgBox Msg$
        snapDivs.Close
        Exit Function
    End If
End If

For x = 1 To 1
    If Len(clpCode(x).Text) > 0 And clpCode(x).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(x).SetFocus
        Exit Function
    End If
Next x

chkBenGrp = True

Exit Function

chkBenGrp_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkBenGrp", "HR_Div", "Cancel")
Resume Next

End Function

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub clpDiv_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmbWDate_Click()
txtWDate.Text = cmbWDate.Text
End Sub

Private Sub cmdBenAll_Click()
Dim rsBenGrp As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String, Msg$, a%
Dim xWhichDate, xFromDate, xToDate, xEmpStatus, xTermReason
Dim xRecCount, I, xEmpCertNo
Dim xCertflag As Boolean
Dim xBenAccflag As Boolean
Dim xCovClassflag As Boolean

On Error GoTo DelErr

SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE (NOT (BM_CERTIFICATE_PREFIX IS NULL OR BM_CERTIFICATE_PREFIX = '' ))"
SQLQ = SQLQ & "ORDER BY BM_BENEFIT_GROUP, BM_DIV ,BM_EMP DESC "
rsBenGrp.Open SQLQ, gdbAdoIhr001, adOpenStatic
xTotRecCount = 0

Msg = "This function will update the Employee's Certificate Number, " & Chr(10)
Msg = Msg & "Benefit Account and Coverage Class on Status/Dates screen. " & Chr(10)
Msg = Msg & "This update will only effect employee records that do not contain any data in them. " & Chr(10)

    'If glbWFC Then
    Msg = Msg & "(Employee Certificate Number = Certificate Number Prefix + Payroll ID)" & Chr(10) & Chr(10)
    'End If
    'Msg = Msg & xTotRecCount & IIf(xTotRecCount = 1, " employee", " employees") & " will be updated." & Chr(10) & Chr(10)
    Msg = Msg & "Are you sure you want to do it?"
    a% = MsgBox(Msg, 36, "Confirm Update")
    If a% <> 6 Then Exit Sub
    
    If Not rsBenGrp.EOF Then
        rsBenGrp.MoveFirst
        xRecCount = rsBenGrp.RecordCount
        I = 0
        MDIMain.panHelp(0).FloodType = 1
        MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(2).Caption = "        "
    End If
    xTotRecCount = 0
    Do While Not rsBenGrp.EOF
        MDIMain.panHelp(0).FloodPercent = (I / xRecCount) * 100
        DoEvents
        I = I + 1
        'Call UpdEmpCertiNum(rsBenGrp("BM_BENEFIT_GROUP"), rsBenGrp("BM_DIV"), rsBenGrp("BM_CERTIFICATE_PREFIX"), rsBenGrp("BM_BENEFIT_ACCOUNT"), rsBenGrp("BM_BENEFIT_CLASS"), 2)
        xWhichDate = "": xFromDate = "": xToDate = "": xEmpStatus = "": xTermReason = ""
        If Not IsNull(rsBenGrp("BM_WHICH_DATE")) Then
            xWhichDate = rsBenGrp("BM_WHICH_DATE")
            If Len(xWhichDate) > 0 Then
                If Not IsNull(rsBenGrp("BM_FROM_DATE")) Then
                    xFromDate = rsBenGrp("BM_FROM_DATE")
                End If
                If Not IsNull(rsBenGrp("BM_TO_DATE")) Then
                    xToDate = rsBenGrp("BM_TO_DATE")
                End If
            End If
        End If
        If Not IsNull(rsBenGrp("BM_EMP")) Then
            xEmpStatus = rsBenGrp("BM_EMP")
        End If
        If Not IsNull(rsBenGrp("BM_TERM_REASON")) Then
            xTermReason = rsBenGrp("BM_TERM_REASON")
        End If
        
        SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1 FROM HREMP  "
        SQLQ = SQLQ & "WHERE ((ED_USER_TEXT1 IS NULL OR LEN(ED_USER_TEXT1) = '' ) OR (ED_USER_TEXT2 IS NULL OR LEN(ED_USER_TEXT2) = '' ) OR (ED_USER_NUM1 IS NULL)) "
        SQLQ = SQLQ & "AND NOT (ED_PAYROLL_ID IS NULL OR ED_PAYROLL_ID = '' ) "
        SQLQ = SQLQ & "AND ED_BENEFIT_GROUP = '" & rsBenGrp("BM_BENEFIT_GROUP") & "' "
        SQLQ = SQLQ & "AND ED_DIV = '" & rsBenGrp("BM_DIV") & "' "
        SQLQ = SQLQ & "AND ED_COUNTRY = 'CANADA' "
        If xWhichDate = "Date of Hire" Then
            If Len(xFromDate) > 0 Then
                SQLQ = SQLQ & "AND ED_DOH >= " & Date_SQL(xFromDate) & " "
            End If
            If Len(xToDate) > 0 Then
                SQLQ = SQLQ & "AND ED_DOH <= " & Date_SQL(xToDate) & " "
            End If
        End If
        If Len(xEmpStatus) > 0 Then
            SQLQ = SQLQ & "AND ED_EMP = '" & xEmpStatus & "' "
        End If
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsTemp.EOF
            xEmpCertNo = ""
            xCertflag = False
            xBenAccflag = False
            xCovClassflag = False
            If Len(rsTemp("ED_PAYROLL_ID")) > 0 Then
                If IsNull(rsTemp("ED_USER_TEXT1")) Then
                    xCertflag = True
                Else
                    If Len(rsTemp("ED_USER_TEXT1")) = 0 Then
                        xCertflag = True
                    End If
                End If
                If xCertflag Then 'Certificate Number
                    xEmpCertNo = Trim(rsBenGrp("BM_CERTIFICATE_PREFIX")) & Trim(rsTemp("ED_PAYROLL_ID"))
                    xEmpCertNo = Right("000000000000" & xEmpCertNo, 12)
                    rsTemp("ED_USER_TEXT1") = xEmpCertNo
                End If
                If IsNull(rsTemp("ED_USER_TEXT2")) Then
                    xBenAccflag = True
                Else
                    If Len(rsTemp("ED_USER_TEXT2")) = 0 Then
                        xBenAccflag = True
                    End If
                End If
                If xBenAccflag Then
                    rsTemp("ED_USER_TEXT2") = rsBenGrp("BM_BENEFIT_CLASS")
                End If
                If IsNull(rsTemp("ED_USER_NUM1")) Then
                    xBenAccflag = True
                End If
                If xBenAccflag Then
                    rsTemp("ED_USER_NUM1") = rsBenGrp("BM_BENEFIT_ACCOUNT")
                End If
                rsTemp.Update
            End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        
        
        rsBenGrp.MoveNext
    Loop

    MsgBox " Update Completed"


    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " "
Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdBenAll", "HR_BENEFITS_GROUP_MATRIX", "UpdateAll")
Resume Next
End Sub

Private Sub cmdBenDiv_Click()
'Dim SQLQ As String, Msg$, a%
'
'On Error GoTo DelErr
'
'xTotRecCount = UpdEmpCertiNum(clpCode(1).Text, clpDiv.Text, txtCertiNum.Text, 1)
'Msg = "This function will update Employee's Certificate Number, Benefit Account" & Chr(10)
'Msg = Msg & "and Coverage Class on Status/Dates screen if they are blank. " & Chr(10)
'Msg = Msg & "for Benfit Group " & clpCode(1).Caption & " " '& Chr(10)
'Msg = Msg & "and " & lStr("Division") & " " & clpDiv.Caption & "." & Chr(10)
'
'
'If xTotRecCount = 0 Then
'    Msg = Msg & Chr(10) & "There is no employee to be updated." & Chr(10)
'    MsgBox Msg
'    Exit Sub
'Else
'    'If glbWFC Then
'    Msg = Msg & "(Employee Certificate Number = Certificate Number Prefix + Payroll ID)" & Chr(10) & Chr(10)
'    'End If
'    Msg = Msg & xTotRecCount & IIf(xTotRecCount = 1, " employee", " employees") & " will be updated." & Chr(10) & Chr(10)
'    Msg = Msg & "Are you sure you want to do it?"
'    a% = MsgBox(Msg, 36, "Confirm Update")
'    If a% <> 6 Then Exit Sub
'
'    Call UpdEmpCertiNum(clpCode(1).Text, clpDiv.Text, txtCertiNum.Text, txtBenAccount, txtCovClass, 2)
'
'    MsgBox " Update Completed"
'End If
'
'Exit Sub
'
'DelErr:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdBenDiv", "HR_BENEFITS_GROUP_MATRIX", "Update")
'Resume Next
End Sub

Private Function UpdEmpCertiNum(xBenGroup, xDiv, xBenPrefix, xBenAccount, xCovClass, xType)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xRecCount As Long
Dim xEmpCertNo As String
Dim XUpdCount As Long


    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1 FROM HREMP WHERE ((ED_USER_TEXT1 IS NULL OR LEN(ED_USER_TEXT1) = '' ) OR (ED_USER_TEXT2 IS NULL OR LEN(ED_USER_TEXT2) = '' ) OR (ED_USER_NUM1 IS NULL)) "
    SQLQ = SQLQ & "AND NOT (ED_PAYROLL_ID IS NULL OR ED_PAYROLL_ID = '' ) "
    SQLQ = SQLQ & "AND ED_BENEFIT_GROUP = '" & xBenGroup & "' "
    SQLQ = SQLQ & "AND ED_DIV = '" & xDiv & "' "
    xRecCount = 0
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        xRecCount = rsEmp.RecordCount
    End If
    If xType = 1 Then 'Get Record Count
        UpdEmpCertiNum = xRecCount
        rsEmp.Close
        Exit Function
    End If
    If xRecCount = 0 Then Exit Function
    XUpdCount = 0
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    MDIMain.panHelp(2).Caption = ""

    Do While Not rsEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (XUpdCount / xRecCount) * 100
        XUpdCount = XUpdCount + 1
        xEmpCertNo = ""
        If Len(rsEmp("ED_PAYROLL_ID")) > 0 Then
            xEmpCertNo = Trim(xBenPrefix) & Trim(rsEmp("ED_PAYROLL_ID"))
            xEmpCertNo = Right("000000000000" & xEmpCertNo, 12)
            rsEmp("ED_USER_TEXT1") = xEmpCertNo
            rsEmp.Update
        End If
        rsEmp.MoveNext
    Loop
    rsEmp.Close
    MDIMain.panHelp(0).FloodType = 0
    
End Function

Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Set_Control("R", Me, rsDATA)


Call modSTUPD(False)  ' reset screen's attributes
cmdClose.SetFocus


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

glbDiv = ""
glbDivDesc = ""

Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim Div As String, SQLQ As String, Msg$, a%

On Error GoTo DelErr


Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh


Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRPROV", "Delete")
Resume Next

End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmdModify_Click()

On Error GoTo Mod_Err

Call modSTUPD(True)
'clpCode(1).Enabled = True
clpCode(1).SetFocus
fglbNewRec% = False
'Data1.Recordset.Edit

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack '08June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()

glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True


Call Set_Control("B", Me)
rsDATA.AddNew


clpCode(1).Enabled = True
clpCode(1).SetFocus


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
Dim xID, ctylist
On Error GoTo OK_Err

If Not chkBenGrp() Then Exit Sub

Call UpdUStats(Me)

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

xID = rsDATA("BM_ID")

Data1.Refresh
Data1.Recordset.Find "BM_ID='" & xID & " '"

fglbNewRec% = False
Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPROV", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub cmdPrint_Click()
Dim RHeading As String, xReport

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = ("Benefit Group Matrix Report")
Me.vbxCrystal.WindowTitle = RHeading

xReport = glbIHRREPORTS & "RzBenGrpMrx.rpt"

Me.vbxCrystal.ReportFileName = xReport

'If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
'Else
'    Me.vbxCrystal.Connect = "PWD=petman;"
'    Me.vbxCrystal.DataFiles(0) = glbIHRDB
'End If

Me.vbxCrystal.Action = 1


End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDiv", "SELECT")

End Sub

Private Sub Form_Load()
Dim SQLQ, I, ctylist, x
glbOnTop = "frmAttendanceGroup"

cmbWDate.AddItem ""
cmbWDate.AddItem "Date of Hire"
cmbWDate.AddItem "Termination"

Data1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX "
SQLQ = SQLQ & " ORDER BY BM_BENEFIT_GROUP,BM_DIV "
Data1.RecordSource = SQLQ
Data1.Refresh

'Data1.LockType = adLockReadOnly

Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT
Call modSTUPD(False)
If Not gSec_BenefitGroupSetup Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If

Call setCaption(lblDiv)
cmdBenDiv.Caption = lStr(cmdBenDiv.Caption)

For I = 1 To 1
    Call setCaption(frmAttendanceGroup.vbxTrueGrid.Columns.Item(I))
Next I

Call Display_Value

Call INI_Controls(Me)

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

cmdModify.Enabled = FT
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '


vbxTrueGrid.Enabled = FT
clpDiv.Enabled = TF
txtCovClass.Enabled = TF
txtBenAccount.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
txtCertiNum.Enabled = TF
cmbWDate.Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
txtComments.Enabled = TF

cmdClose.Enabled = FT
'cmdPrint.Enabled = FT           '
        
End Sub



Private Sub txtBenAccount_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtBenAccount_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtCertiNum_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtCovClass_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtWDate_Change()
    cmbWDate.Text = txtWDate.Text
End Sub

Private Sub vbxTrueGrid_DblClick()
    
If Not Me.vbxTrueGrid.EditActive Then
    glbDiv = Data1.Recordset("DIV")
    glbDivDesc = Data1.Recordset("Division_Name")
    Unload Me
Else
    MsgBox "Save/cancel changes first"
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
        
        SQLQ = "select * from HR_BENEFITS_GROUP_MATRIX WHERE " & glbSeleDiv
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    If Me.vbxTrueGrid.EditActive Then
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


Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'''Sam add July 02 * Remove ADO
Call Display_Value
End Sub
''' Sam add July 2002 * Remove ADO
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        'SQLQ = "select * from HR_DIVISION WHERE DIV='" & Data1.Recordset!Div & "'" & " order by Division_Name"
        SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX "
        SQLQ = SQLQ & "WHERE BM_ID='" & Data1.Recordset!BM_ID & "'"
        SQLQ = SQLQ & " ORDER BY BM_BENEFIT_GROUP,BM_DIV "
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
        Call Set_Control("R", Me, rsDATA)
    End If
    
End Sub

Private Sub IsNew74()
    Dim SQLQ
 
    SQLQ = "select * from HR_DIVISION WHERE DIV='" & Data1.Recordset!Div & "'" & " order by Division_Name"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub

Function GetBonusRptDesc(TablKey)
    Dim rsTABL As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT * FROM WFC_Bonus_Loc_Department WHERE Dept_No = '" & TablKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTABL.EOF And rsTABL.BOF Then
        GetBonusRptDesc = ""
    Else
        GetBonusRptDesc = rsTABL("Dept_Name")
    End If
    rsTABL.Close
End Function

