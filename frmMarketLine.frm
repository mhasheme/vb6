VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMarketLine 
   Caption         =   "Market Line"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   7185
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "ML_CODE"
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   10
      Tag             =   "01-Market Line"
      Top             =   4200
      Width           =   1065
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "ML_DESC"
      Height          =   285
      Left            =   1320
      MaxLength       =   25
      TabIndex        =   0
      Tag             =   "01-Description of Code"
      Top             =   4200
      Width           =   3615
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5640
      Top             =   4440
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
      TabIndex        =   1
      Top             =   4875
      Width           =   7185
      _Version        =   65536
      _ExtentX        =   12674
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
         Left            =   5895
         TabIndex        =   9
         Tag             =   "Print Division Listing"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
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
         Left            =   5070
         TabIndex        =   8
         Tag             =   "Delete Division listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
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
         Left            =   4260
         TabIndex        =   7
         Tag             =   "Create a new Division"
         Top             =   105
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
         Left            =   3360
         TabIndex        =   6
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
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
         Left            =   2520
         TabIndex        =   5
         Tag             =   "Save changes made"
         Top             =   105
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
         Left            =   1680
         TabIndex        =   4
         Tag             =   "Edit the information "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
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
         Left            =   855
         TabIndex        =   3
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
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
         Left            =   15
         TabIndex        =   2
         Tag             =   "Select this Division"
         Top             =   105
         Width           =   735
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
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmMarketLine.frx":0000
      Height          =   4005
      Left            =   120
      OleObjectBlob   =   "frmMarketLine.frx":0014
      TabIndex        =   11
      Tag             =   "Division Listings"
      Top             =   120
      Width           =   6975
   End
End
Attribute VB_Name = "frmMarketLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fglbNewRec% ' new record
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO

Private Sub cmdCancel_Click()
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Set_Control("R", Me, rsDATA)


Call modSTUPD(False)  ' reset screen's attributes
cmdClose.SetFocus


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "Market Line", "Cancel")
Resume Next
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim Div As String, SQLQ As String, Msg$, a%
Dim snapEEDivs As New ADODB.Recordset

On Error GoTo DelErr

If Len(txtCode) < 1 Then Exit Sub

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

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call modSTUPD(True)

txtName.SetFocus


Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack
End Sub

Private Sub cmdNew_Click()
glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True
txtCode.Text = ""
txtName.Text = ""
Call Set_Control("B", Me)
rsDATA.AddNew


txtCode.Enabled = True
txtCode.SetFocus


Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "Market Line", "AddNew")
Resume Next

End Sub

Private Sub cmdOK_Click()
Dim DivCode
On Error GoTo OK_Err

If Not chkMarketLine() Then Exit Sub

DivCode = txtCode

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
Data1.Recordset.Find "ML_CODE='" & txtCode & " '"


fglbNewRec% = False
Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "Market Line", "Update")
Resume Next
Unload Me

End Sub

Private Sub cmdPrint_Click()
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    RHeading = "Market Line Listing Report"
    Me.vbxCrystal.WindowTitle = RHeading
    Me.vbxCrystal.BoundReportHeading = RHeading

    xReport = glbIHRREPORTS & "WFCmarket.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    Me.vbxCrystal.Connect = RptODBC_SQL
    Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdSelect_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim SQLQ
glbOnTop = "FRMMARKETLINE"
Data1.ConnectionString = glbAdoIHRDB
SQLQ = "select * from WFC_MARKETLINE_DESC order by ML_CODE"
Data1.RecordSource = SQLQ
Data1.Refresh
Data1.LockType = adLockReadOnly

Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT
Call modSTUPD(False)

Call INI_Controls(Me)

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
txtName.Enabled = TF            '
cmdClose.Enabled = FT           '
cmdSelect.Enabled = False          '
cmdPrint.Enabled = FT           '

End Sub

Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        Exit Sub
    End If
  
    SQLQ = "select * from WFC_MARKETLINE_DESC WHERE ML_CODE='" & Data1.Recordset!ML_CODE & "'" & " order by ML_CODE"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "select * from WFC_MARKETLINE_DESC "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub

Private Function chkMarketLine()
Dim xCode As String, SQLQ As String, Msg$
Dim snapMarket As New ADODB.Recordset

chkMarketLine = False
On Error GoTo chkMarketLine_Err

If Len(txtCode) < 1 Then
    MsgBox "Market Line Code is a required field"
    txtCode.SetFocus
    Exit Function
End If

If Len(txtName) < 1 Then
    MsgBox "Description is a required field"
    txtName.SetFocus
    Exit Function
End If

If fglbNewRec Then
    xCode = txtCode
    SQLQ = "SELECT * from WFC_MARKETLINE_DESC "
    SQLQ = SQLQ & "WHERE ML_CODE = '" & xCode & "'"
    
    If snapMarket.State <> 0 Then snapMarket.Close
    snapMarket.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapMarket.BOF And snapMarket.EOF Then
        snapMarket.Close
    Else
        Msg$ = "This Market Line Code already exists"
        MsgBox Msg$
        snapMarket.Close
        Exit Function
    End If
End If

chkMarketLine = True

Exit Function

chkMarketLine_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkMarketLine", "HR_Div", "Cancel")
Resume Next

End Function
