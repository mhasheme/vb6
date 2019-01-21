VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmFORMALED 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Formal Education"
   ClientHeight    =   8490
   ClientLeft      =   240
   ClientTop       =   1515
   ClientWidth     =   11880
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
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkCompPaid 
      Height          =   255
      Left            =   2370
      TabIndex        =   6
      Tag             =   "40-Company Paid -y/n"
      Top             =   5190
      Width           =   315
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   270
      Left            =   2700
      TabIndex        =   25
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EU_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   7200
      MaxLength       =   25
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin INFOHR_Controls.DateLookup dlpYear 
      DataField       =   "EU_YEAR"
      Height          =   285
      Left            =   2040
      TabIndex        =   4
      Tag             =   "40-Completed on - enter date"
      Top             =   4440
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EU_DEGREE"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   5
      Tag             =   "00-Degree Obtained - Code"
      Top             =   4800
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EUDE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EU_MINOR"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   2
      Tag             =   "00-Minor Study - Code"
      Top             =   3720
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EUMJ"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EU_MAJOR"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   1
      Tag             =   "00-Major Study - Code"
      Top             =   3360
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EUMJ"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EU_SCHOOL"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   0
      Tag             =   "01-School - Code"
      Top             =   3000
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EUSC"
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   22
      Top             =   7830
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8400
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
         _ExtentX        =   3625
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   7380
         Top             =   95
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
   Begin VB.CheckBox chkCompleted 
      Height          =   255
      Left            =   2370
      TabIndex        =   3
      Tag             =   "40-Completed -y/n"
      Top             =   4110
      Value           =   1  'Checked
      Width           =   315
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EU_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   3720
      MaxLength       =   25
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "EU_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5520
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
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
         Left            =   7080
         TabIndex        =   27
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
         Left            =   1320
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
         Left            =   2880
         TabIndex        =   10
         Top             =   135
         Width           =   720
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feformed.frx":0000
      Height          =   2115
      Left            =   120
      OleObjectBlob   =   "feformed.frx":0014
      TabIndex        =   24
      Top             =   720
      Width           =   9015
   End
   Begin VB.Label lblCompanyPaid 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      DataField       =   "EU_COMPANY_PAID"
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
      Left            =   3120
      TabIndex        =   29
      Top             =   5280
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Company Paid"
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
      Left            =   330
      TabIndex        =   28
      Top             =   5220
      Width           =   1020
   End
   Begin VB.Image imgNoSec 
      Height          =   240
      Left            =   2280
      Picture         =   "feformed.frx":4214
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgSec 
      Height          =   240
      Left            =   2280
      Picture         =   "feformed.frx":435E
      Top             =   5880
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblImport 
      Caption         =   "Formal Education"
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
      Left            =   330
      TabIndex        =   26
      Top             =   5880
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Completed"
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
      Left            =   330
      TabIndex        =   21
      Top             =   4175
      Width           =   750
   End
   Begin VB.Label lblCompleted 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Y"
      DataField       =   "EU_COMP"
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
      Left            =   3360
      TabIndex        =   20
      Top             =   4170
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Degree Obtained"
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
      Index           =   5
      Left            =   330
      TabIndex        =   19
      Top             =   4890
      Width           =   1455
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Completed on     "
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
      TabIndex        =   18
      Top             =   4530
      Width           =   1200
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Minor Study    "
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
      Top             =   3810
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Major Study    "
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
      TabIndex        =   16
      Top             =   3450
      Width           =   1020
   End
   Begin VB.Label lblTitle 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "School"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   330
      TabIndex        =   15
      Top             =   3090
      Width           =   600
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EU_EMPNBR"
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
      Left            =   1695
      TabIndex        =   13
      Top             =   6600
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EU_COMPNO"
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
      TabIndex        =   14
      Top             =   6600
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmFORMALED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO

Private Sub chkCompleted_Click()
   
If chkCompleted.Value = 1 Then
    lblCompleted.Caption = "Y"
    dlpYear.Visible = True
    clpCode(4).Visible = True
Else
    lblCompleted.Caption = "N"
    dlpYear.Visible = False
    clpCode(4).Visible = False
End If
lblTitle(4).Visible = dlpYear.Visible
lblTitle(5).Visible = dlpYear.Visible
End Sub

Private Sub chkCompleted_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Function chkFormalEd()
Dim oCode As String, OCodeD As String

chkFormalEd = False

If Len(clpCode(1).Text) <= 0 Then
    MsgBox "Valid school code is a required field"
    clpCode(1).SetFocus
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "School code must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(2).Text) > 0 Then
    If clpCode(2).Caption = "Unassigned" Then
        MsgBox "Major Study code must be valid"
        clpCode(2).SetFocus
        Exit Function
    End If
End If

If Len(clpCode(3).Text) > 0 Then
    If clpCode(3).Caption = "Unassigned" Then
        MsgBox "Minor Study code must be valid"
        clpCode(3).SetFocus
        Exit Function
    End If
End If

If chkCompleted.Value = 1 And Len(clpCode(4).Text) > 0 Then
    If clpCode(4).Caption = "Unassigned" Then
        MsgBox "Degree is a required field"
        clpCode(4).SetFocus
        Exit Function
    End If
End If

If Len(dlpYear.Text) > 0 Then
    If Not IsDate(dlpYear.Text) Then
        MsgBox "Not a valid date"
        dlpYear.SetFocus
        Exit Function
    End If
End If

If chkCompleted.Value = 1 Then
    lblCompleted.Caption = "Y"
Else
    lblCompleted.Caption = "N"
End If

If chkCompPaid.Value = 1 Then
    lblCompanyPaid.Caption = "Y"
Else
    lblCompanyPaid.Caption = "N"
End If

chkFormalEd = True

End Function

Sub cmdCancel_Click()
Dim X
On Error GoTo Can_Err


rsDATA.CancelUpdate
Call Display_Value

fglbNew = False
'Call ST_UPD_MODE(True)  ' reset screen's attributes
Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREDU", "Cancel")
Resume Next

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMFORMALED" Then glbOnTop = ""

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

If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "Delete from Term_HRDOC_HREDU where EU_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq & " and EU_DOCKEY=" & glbDocKey & " "
    End If
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    If gsAttachment_DB Then
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_HREDU where EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR = " & glbLEE_ID & " and EU_DOCKEY=" & glbDocKey & " "
    End If
    Data1.Refresh
End If
'If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
'End If

fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREDU", "Delete")
Resume Next
Unload Me

End Sub




Sub cmdNew_Click()
Dim X%

fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

clpCode(1).Enabled = True
clpCode(1).SetFocus
On Error GoTo AddN_Err

If gsAttachment_DB Then
    lblImport.Visible = True 'False
    imgSec.Visible = False
    imgNoSec.Visible = True 'False
    cmdImport.Visible = True 'False
End If

'Data1.Recordset.AddNew
''' Sam add July 2002 * Remove ADO

Call Set_Control("B", Me)

rsDATA.AddNew

chkCompleted.Value = 0
chkCompPaid.Value = 0
lblCNum.Caption = "001"

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HREDU", "Add")

MsgBox Err.Description
Resume Next

End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim X
Dim xID As Long

On Error GoTo Add_Err

If Not chkFormalEd() Then Exit Sub

Call UpdUStats(Me) ' update user's stats (who did it and when)
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

If Not (rsDATA.EOF And rsDATA.BOF) Then
    xID = rsDATA("EU_ID")
End If

Data1.Refresh

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

fglbNew = False

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

If NextFormIF("Formal Education") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREDU", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Formal Education"
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

RHeading = lblEEName & "'s Formal Education"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Private Sub chkCompPaid_Click()
If chkCompPaid.Value = 1 Then
    lblCompanyPaid.Caption = "Y"
Else
    lblCompanyPaid.Caption = "N"
End If
End Sub

Private Sub chkCompPaid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdImport_Click()
    glbDocNewRecord = fglbNew
    glbDocName = "FormalEdu"
    If fglbNew Then
        glbDocKey = 0
    Else
        glbDocKey = rsDATA("EU_ID")
    End If
    frmInAttachment.Show 1
    DoEvents
    Call DispimgIcon(Me, "frmFORMALED")
End Sub

'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_Formal ed(hredu)S", "SELECT")

End Sub



Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

'SQLQ = "Select * FROM HREDU"       'Lucy July 4, 2000
'SQLQ = SQLQ & " WHERE (HREDU.EU_EMPNBR = " & glbLEE_ID & ") "
'SQLQ = SQLQ & "ORDER BY EU_SCHOOL"

If glbtermopen Then         'Lucy July 4, 2000
    SQLQ = "Select * from Term_EDU"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY EU_SCHOOL"
Else
    SQLQ = "Select * from HREDU"
    SQLQ = SQLQ & " where EU_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY EU_SCHOOL"
End If

Data1.RecordSource = SQLQ
Data1.Refresh
EERetrieve = True






Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DEPRetrieve", "HREDU", "SELECT")
Resume Next

Exit Function

End Function

Private Sub Form_Activate()
    glbOnTop = "FRMFORMALED"
    Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMFORMALED"
End Sub

Private Sub Form_Load()
    Dim Answer, DefVal, Msg, Title  '  variables.
    Dim RFound As Integer, X% ' records found
    
    glbOnTop = "FRMFORMALED"
    If glbtermopen Then         'Lucy July 4, 2000
        Data1.ConnectionString = glbAdoIHRAUDIT
        Data1.RecordSource = "Term_EDU"
    Else
        Data1.ConnectionString = glbAdoIHRDB
        Data1.RecordSource = "HREDU"
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
    
    Call Display_Value
    
    Screen.MousePointer = HOURGLASS
    
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        frmFORMALED.Caption = "Formal Education - " & Left$(glbLEE_SName, 5)
        frmFORMALED.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    
    lblEENum.Caption = ShowEmpnbr(lblEEID)
    
    Call ST_UPD_MODE(True)
    Call SET_UP_MODE
    
    ''If Not gSec_Upd_Formal_Education Then
    '    Call ST_UPD_MODE(False)
    '    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    '        cmdModify.Enabled = False
    '    Else
    '        cmdModify.Enabled = True
    '    End If
    'Else
    ''    Call ST_UPD_MODE(False)             '
    ''    cmdModify.Enabled = False
    ''    cmdNew.Enabled = False
    ''    cmdDelete.Enabled = False
    ''End If
    
    
    chkCompleted.Enabled = False
    chkCompPaid.Enabled = False
    clpCode(1).Enabled = False
    clpCode(2).Enabled = False
    clpCode(3).Enabled = False
    clpCode(4).Enabled = False
    dlpYear.Enabled = False
        
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
'    Set frmFORMALED = Nothing 'carmen may 00
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

fUPMode = TF    ' update mode
'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'vbxTrueGrid.Enabled = FT
chkCompleted.Enabled = TF
chkCompPaid.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
clpCode(3).Enabled = TF
clpCode(4).Enabled = TF
dlpYear.Enabled = TF

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If

glbDocName = "FormalEdu"
If gsAttachment_DB Then
    If Not (Data1.Recordset.BOF And Data1.Recordset.EOF) Then
        If rsDATA.RecordCount > 0 Then
            If Not IsNull(rsDATA("EU_DOCKEY")) Then
                glbDocKey = rsDATA("EU_DOCKEY")
            Else
                glbDocKey = 0
            End If
        Else
            If Not IsNull(Data1.Recordset("EU_DOCKEY")) Then
                glbDocKey = Data1.Recordset("EU_DOCKEY")
            Else
                glbDocKey = 0
            End If
        End If
    End If
    
    Call DispimgIcon(Me, "frmFORMALED")
    If gSec_Upd_Formal_Education Then
        If Data1.Recordset.BOF And Data1.Recordset.EOF Then
            cmdImport.Visible = False
        Else
            cmdImport.Visible = True
        End If
    End If
End If

End Sub

Private Sub lblCompanyPaid_Change()
    If lblCompanyPaid.Caption = "Y" Then
        chkCompPaid.Value = 1
    Else
        chkCompPaid.Value = 0
    End If
End Sub


Private Sub lblCompleted_Change()
If lblCompleted.Caption = "Y" Then
    chkCompleted.Value = 1
Else
    chkCompleted.Value = 0
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
        
        If glbtermopen Then         'Lucy July 4, 2000
            SQLQ = "Select * from Term_EDU"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HREDU"
            SQLQ = SQLQ & " where EU_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

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

''' Sam add July 2002 * Remove ADO
Sub Display_Value()
    Dim SQLQ

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
       
    If glbtermopen Then
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
Else
    If glbtermopen Then
        SQLQ = "Select * from Term_EDU"
        SQLQ = SQLQ & " WHERE EU_ID = " & Data1.Recordset!EU_ID
        SQLQ = SQLQ & " ORDER BY EU_SCHOOL"
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select * from HREDU"
        SQLQ = SQLQ & " where EU_ID = " & Data1.Recordset!EU_ID
        SQLQ = SQLQ & " ORDER BY EU_SCHOOL"
        
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        
        If glbOracle Then   'If used on SQL version then it gives "object in a zombie state error"
            rsDATA.CursorLocation = adUseServer
        End If
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
UpdateRight = gSec_Upd_Formal_Education
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
    frmFORMALED.Caption = "Formal Education - " & Left$(glbLEE_SName, 5)
    frmFORMALED.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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

Private Sub imgSec_Click()
    Dim SQLQ
    SQLQ = getSQL("frmFORMALED")
    Call FillMemoFile(SQLQ, "FormalEdu")
End Sub


