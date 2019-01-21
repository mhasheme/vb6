VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEPayTran 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Payroll Transactions"
   ClientHeight    =   7815
   ClientLeft      =   105
   ClientTop       =   975
   ClientWidth     =   10095
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
   ScaleHeight     =   7815
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtIndicator 
      Appearance      =   0  'Flat
      DataField       =   "PT_INDICATOR"
      Height          =   285
      Left            =   4080
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox comIndicator 
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
      Left            =   2595
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Tag             =   "10-Type of Employee "
      Top             =   3240
      Width           =   1455
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmPayTranImp.frx":0000
      Height          =   2235
      Left            =   120
      OleObjectBlob   =   "frmPayTranImp.frx":0014
      TabIndex        =   0
      Tag             =   "Listing of Other Earnings"
      Top             =   570
      Width           =   8805
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "PT_PAYEND"
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   4
      Tag             =   "41-Ending date of period"
      Top             =   3900
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "PT_PAYSTART"
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   3
      Tag             =   "41-Starting date of period"
      Top             =   3570
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PT_PAYCODE"
      Height          =   285
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Tag             =   "01-Pay Code"
      Top             =   2940
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "PYTC"
      MaxLength       =   10
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   21
      Top             =   7395
      Width           =   10095
      _Version        =   65536
      _ExtentX        =   17806
      _ExtentY        =   741
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
      Begin MSAdodcLib.Adodc DATA1 
         Height          =   405
         Left            =   9600
         Top             =   0
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   714
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
         Left            =   8250
         Top             =   150
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
      DataField       =   "PT_LDATE"
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
      Left            =   4365
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6690
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PT_LTIME"
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
      Left            =   6000
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PT_LUSER"
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
      Left            =   7635
      MaxLength       =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6690
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   10095
      _Version        =   65536
      _ExtentX        =   17806
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
         TabIndex        =   24
         Top             =   120
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
         Left            =   3120
         TabIndex        =   12
         Top             =   120
         Width           =   1740
      End
   End
   Begin MSMask.MaskEdBox medAmount 
      DataField       =   "PT_DOLLARAMT"
      Height          =   285
      Left            =   2595
      TabIndex        =   5
      Tag             =   "20- Actual amount of earnings during this period"
      Top             =   4215
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
      Format          =   "##0.00;(##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medActual 
      Height          =   285
      Left            =   2595
      TabIndex        =   6
      Tag             =   "20- Actual amount of earnings during this period"
      Top             =   4530
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      ClipMode        =   1
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
      Format          =   "##0.00;(##0.00)"
      PromptChar      =   "_"
   End
   Begin INFOHR_Controls.CodeLookup clpVadim2 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Top             =   4850
      Visible         =   0   'False
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDV2"
   End
   Begin VB.Label lblVadim21 
      AutoSize        =   -1  'True
      Caption         =   "Vadim Field 2"
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
      Left            =   210
      TabIndex        =   26
      Top             =   4845
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Actual"
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
      Index           =   6
      Left            =   210
      TabIndex        =   25
      Top             =   4560
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Earning/Deduction Indicator"
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
      Left            =   210
      TabIndex        =   22
      Top             =   3240
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "To Date"
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
      Left            =   210
      TabIndex        =   20
      Top             =   3945
      Width           =   705
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "From Date"
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
      Left            =   210
      TabIndex        =   19
      Top             =   3615
      Width           =   885
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   210
      TabIndex        =   18
      Top             =   4245
      Width           =   1365
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Code"
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
      Left            =   210
      TabIndex        =   17
      Top             =   2955
      Width           =   1455
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "PT_EMPNBR"
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
      Left            =   8310
      TabIndex        =   15
      Top             =   6315
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "PT_COMPNO"
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
      Left            =   7470
      TabIndex        =   16
      Top             =   6315
      Visible         =   0   'False
      Width           =   570
   End
End
Attribute VB_Name = "frmEPayTran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim OADOLLAR, OEARN, OCOEFLAG, Actn
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim fglbNew As Integer
Dim FRS As ADODB.Recordset


Private Function chkEOTHERE()
Dim SQLQ As String, Msg As String
Dim dd&

chkEOTHERE = False

On Error GoTo chkEOTHERE_Err

If Len(clpCode(1).Text) < 1 Then
    MsgBox "Earnings code is a required field"
    clpCode(1).SetFocus
    Exit Function
Else
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "Earnings code must be valid"
        clpCode(1).SetFocus
        Exit Function
    End If
End If
If Len(Trim(medAmount)) > 0 Then      'laura jan 12, 1998
    If Not IsNumeric(medAmount) Then
        MsgBox "Amount is invalid"
        medAmount.SetFocus
        Exit Function
    End If
Else
    medAmount = 0
End If

If Len(dlpDate(0).Text) < 1 Then
    MsgBox "From Date is required field"
    dlpDate(0).SetFocus
    Exit Function
End If

If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "From Date is not a valid date"
        dlpDate(0).SetFocus
        Exit Function
    End If
End If

If Len(dlpDate(1).Text) >= 1 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "To Date is not a valid date"
        dlpDate(1).SetFocus
        Exit Function
    End If
Else
    MsgBox "To Date is required field"
    dlpDate(1).SetFocus
    Exit Function
End If

dd& = DateDiff("d", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text))

If dd& < 0 Then
    MsgBox "To Date cannot precede From Date"
    dlpDate(1).SetFocus
    Exit Function
End If

''Add by Frank Dec 18,2001
'If glbCompSerial = "S/N - 2214W" Then 'Casey House
'    If Not chkCOEFlag Then
'        MsgBox "Cost of Employment must be On"
'        chkCOEFlag.SetFocus
'        Exit Function
'    End If
'End If

chkEOTHERE = True

Exit Function

chkEOTHERE_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkOThern", "HR_PAYROLL_TRANSACTION", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err
fglbNew = False
rsDATA.CancelUpdate
Call Display_Value

Exit Sub
Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_PAYROLL_TRANSACTION", "Cancel")
Resume Next

End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "frmePayTran" Then glbOnTop = ""

End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, x
Dim xID

If DATA1.Recordset.BOF And DATA1.Recordset.EOF Then
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

If vbxTrueGrid.SelBookmarks.count = 0 Then vbxTrueGrid.SelBookmarks.Add DATA1.Recordset.Bookmark
For x = 0 To vbxTrueGrid.SelBookmarks.count - 1
    DATA1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(x)
    xID = DATA1.Recordset("PT_ID")
    If glbtermopen Then
        gdbAdoIhr001X.Execute "DELETE FROM Term_PAYROLL_TRANSACTION WHERE PT_ID=" & xID
    Else
        gdbAdoIhr001.Execute "DELETE FROM HR_PAYROLL_TRANSACTION WHERE PT_ID=" & xID
    End If
Next
DATA1.Refresh

If DATA1.Recordset.EOF And DATA1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False
Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_PAYROLL_TRANSACTION", "Delete")
Call RollBack '26July99 js

End Sub

Sub cmdModify_Click()

On Error GoTo Mod_Err

Actn = "M"
OADOLLAR = medAmount
OEARN = clpCode(1).Text

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_PAYROLL_TRANSACTION", "Modify")
Call RollBack

End Sub

Sub cmdNew_Click()
Dim SQLQ As String
fglbNew = True

Call SET_UP_MODE

On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

Actn = "A"

If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"
clpCode(1).SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_PAYROLL_TRANSACTION", "Add")
Resume Next

End Sub

Sub cmdOK_Click()
Dim xID As Long
On Error GoTo Add_Err
If Not chkEOTHERE Then Exit Sub

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
DATA1.Refresh
fglbNew = False
Call SET_UP_MODE

Exit Sub

Add_Err:
If Err = 3022 Then
    DATA1.Recordset.CancelUpdate
    DATA1.Refresh
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PAYROLL_TRANSACTION", "Update")
Resume Next
Unload Me

End Sub
Private Function Round53(xNumb)
Dim xInteger, xDecimal, xDecTmp
    xInteger = Int(xNumb * 1000)
    xDecimal = xNumb * 1000 - xInteger
    xDecTmp = 0
    If xDecimal >= 0 And xDecimal < 0.5 Then
        xDecTmp = 0
    End If
    If xDecimal >= 0.5 Then
        xDecTmp = 1
    End If
    Round53 = (xInteger + xDecTmp) / 1000
End Function

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Payroll Transactions"
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

RHeading = lblEEName & "'s Payroll Transactions"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_PAYROLL_TRANSACTION"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY PT_PAYEND DESC,PT_PAYCODE "
Else
    SQLQ = "Select * from HR_PAYROLL_TRANSACTION"
    SQLQ = SQLQ & " where PT_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY PT_PAYEND DESC,PT_PAYCODE "
End If

DATA1.RecordSource = SQLQ
DATA1.Refresh
Set FRS = DATA1.Recordset.Clone

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Earnings", "HR_PAYROLL_TRANSACTION", "SELECT")
Resume Next

Exit Function

End Function



Private Sub comIndicator_LostFocus()
If comIndicator.ListIndex = 0 Then
    txtIndicator.Text = "D"
End If
If comIndicator.ListIndex = 1 Then
    txtIndicator.Text = "E"
End If
If glbCompSerial = "S/N - 2431W" Then 'BACI Ticket #22736 Franks 11/01/2012
    If comIndicator.ListIndex = 2 Then
        txtIndicator.Text = "H"
    End If
Else
'If glbWFC Then 'Ticket #21220 Franks 11/29/2011
    If comIndicator.ListIndex = 2 Then
        txtIndicator.Text = "M"
    End If
'End If
End If
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
    Me.cmdModify_Click
    glbOnTop = "frmePayTran"
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "frmePayTran"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "frmePayTran"

If glbtermopen Then
    DATA1.ConnectionString = glbAdoIHRAUDIT
Else
    DATA1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT

'Ticket #21120 Franks 11/22/2011 - begin
If glbWFC Then
    medActual.DataField = "PT_ACTUAL"
    medActual.Visible = True
    lblTitle(6).Visible = True
    clpVadim2.DataField = "PT_VADIM2"
    clpVadim2.Visible = True
    lblVadim21.Caption = lStr("Vadim Field 2")
    lblVadim21.Visible = True
    vbxTrueGrid.Columns(6).Caption = lStr("Vadim Field 2")
Else
    vbxTrueGrid.Columns(5).Visible = False
    vbxTrueGrid.Columns(6).Visible = False
End If
'Ticket #21120 Franks 11/22/2011 - end

If glbCompSerial = "S/N - 2431W" Then 'BACI Ticket #21638 Franks 10/01/2012
    medActual.DataField = "PT_ACTUAL"
    medActual.Visible = True
    lblTitle(6).Visible = True
    vbxTrueGrid.Columns(5).Visible = True
    vbxTrueGrid.Columns(5).NumberFormat = "Standard" 'Ticket #22736 Franks 11/01/2012
End If

Call ComEType

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

vbxTrueGrid.FetchRowStyle = True
vbxTrueGrid.MarqueeStyle = 3

If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Payroll Transactions - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call ST_UPD_MODE(True)
Call Display_Value

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmEPayTran = Nothing
    Call NextForm
End Sub

Private Sub medAmount_GotFocus()
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

'chkCOEFlag.Enabled = TF
medAmount.Enabled = TF
clpCode(1).Enabled = TF
dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
comIndicator.Enabled = TF

If DATA1.Recordset.EOF Or DATA1.Recordset.BOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
End Sub

Private Sub medAmount_LostFocus()
    If Len(Trim(medAmount)) = 0 Then medAmount = 0
End Sub


Private Sub txtIndicator_Change()
comIndicator.ListIndex = -1

If txtIndicator = "E" Then
    comIndicator.ListIndex = 1
End If
If txtIndicator = "D" Then
    comIndicator.ListIndex = 0
End If
If glbCompSerial = "S/N - 2431W" Then 'BACI Ticket #22736 Franks 11/01/2012
    If txtIndicator = "H" Then
        comIndicator.ListIndex = 2
    End If
Else
'If glbWFC Then 'Ticket #21220 Franks 11/29/2011
    If txtIndicator = "M" Then
        comIndicator.ListIndex = 2
    End If
'End If
End If
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub



Private Sub vbxTrueGrid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid60.StyleDisp)
    On Error GoTo EH

    FRS.Requery
    FRS.Bookmark = Bookmark

EH:
    Exit Sub
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
            SQLQ = "Select * from Term_PAYROLL_TRANSACTION"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HR_PAYROLL_TRANSACTION"
            SQLQ = SQLQ & " where PT_EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        DATA1.RecordSource = SQLQ
        DATA1.Refresh
        Set FRS = DATA1.Recordset.Clone
        vbxTrueGrid.FetchRowStyle = True

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value


Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_PAYROLL_TRANSACTION", "Add")
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
    If DATA1.Recordset.EOF Or DATA1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open DATA1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open DATA1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Me.cmdModify_Click
        Exit Sub
    End If
    
If glbtermopen Then
    SQLQ = "Select * from Term_PAYROLL_TRANSACTION"
    SQLQ = SQLQ & " WHERE PT_ID = " & DATA1.Recordset!PT_ID
    SQLQ = SQLQ & " ORDER BY PT_PAYEND, PT_PAYCODE"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
Else
    SQLQ = "Select * from HR_PAYROLL_TRANSACTION"
    SQLQ = SQLQ & " where PT_ID = " & DATA1.Recordset!PT_ID
    SQLQ = SQLQ & " ORDER BY PT_PAYEND, PT_PAYCODE"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If


    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call SET_UP_MODE
    Me.cmdModify_Click
    Call Set_Control("R", Me, rsDATA)

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
UpdateRight = gSec_Upd_PayrollTrans
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
    frmEPayTran.Caption = "Payroll Transactions - " & Left$(glbLEE_SName, 5)
    frmEPayTran.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
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


Private Sub ComEType()
comIndicator.Clear
comIndicator.AddItem "E - Earning"
comIndicator.AddItem "D - Deduction"
If glbCompSerial = "S/N - 2431W" Then 'BACI Ticket #22736 Franks 11/01/2012
    comIndicator.AddItem "H - Hours"
Else
'If glbWFC Then 'Ticket #21220 Franks 11/29/2011
    comIndicator.AddItem "M - Memo"
'End If
End If
End Sub



