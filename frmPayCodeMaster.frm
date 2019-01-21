VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#60.0#0"; "IHRCTRLS.OCX"
Begin VB.Form frmPayCodeMaster 
   Caption         =   "Pay Period Master"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Caption         =   "Payroll ->HR"
      Height          =   225
      Left            =   2610
      TabIndex        =   26
      Top             =   5310
      Width           =   1485
   End
   Begin VB.CheckBox chkTransferHR 
      Caption         =   "HR -> Payroll"
      Height          =   225
      Left            =   510
      TabIndex        =   25
      Top             =   5280
      Width           =   1485
   End
   Begin VB.OptionButton Option5 
      Caption         =   "Daily"
      Height          =   225
      Left            =   7470
      TabIndex        =   24
      Top             =   4770
      Width           =   1095
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Weekly"
      Height          =   225
      Left            =   6300
      TabIndex        =   23
      Top             =   4770
      Width           =   1095
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Pay Period"
      Height          =   225
      Left            =   5040
      TabIndex        =   22
      Top             =   4770
      Width           =   1095
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Monthly"
      Height          =   225
      Left            =   3780
      TabIndex        =   21
      Top             =   4770
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Annual"
      Height          =   225
      Left            =   2490
      TabIndex        =   20
      Top             =   4740
      Width           =   1095
   End
   Begin VB.ComboBox cmbType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2220
      TabIndex        =   16
      Tag             =   "11-Choose Type of Payroll Matrix"
      Top             =   3540
      Width           =   1515
   End
   Begin VB.ComboBox comPayTypeIDCode 
      DataField       =   "PAY_TYPE_ID_CODE"
      Height          =   315
      Left            =   2250
      TabIndex        =   13
      Top             =   3000
      Width           =   1395
   End
   Begin VB.ComboBox comPayTypeCode 
      DataField       =   "PAY_TYPE_CODE"
      Height          =   315
      Left            =   2250
      TabIndex        =   12
      Top             =   2670
      Width           =   1395
   End
   Begin VB.TextBox txtPayCodeName 
      Appearance      =   0  'Flat
      DataField       =   "PAY_CODE_NAME"
      Height          =   285
      Left            =   2280
      TabIndex        =   9
      Tag             =   "61-Year"
      Top             =   2370
      Width           =   3885
   End
   Begin VB.TextBox txtPayCode 
      Appearance      =   0  'Flat
      DataField       =   "PAY_CODE"
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "61-Year"
      Top             =   2070
      Width           =   1215
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   7
      Top             =   7170
      Width           =   11865
      _Version        =   65536
      _ExtentX        =   20929
      _ExtentY        =   979
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
         Left            =   3570
         TabIndex        =   6
         Tag             =   "Print Listing "
         Top             =   60
         Width           =   855
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
         Left            =   2610
         TabIndex        =   5
         Tag             =   "Cancel the changes made"
         Top             =   60
         Width           =   855
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
         Left            =   1770
         TabIndex        =   4
         Tag             =   "Save the changes made"
         Top             =   60
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
         Left            =   900
         TabIndex        =   3
         Tag             =   "Edit the information on this screen"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
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
         Left            =   120
         TabIndex        =   2
         Tag             =   "Close and exit this screen"
         Top             =   60
         Width           =   735
      End
      Begin MSAdodcLib.Adodc datPP 
         Height          =   330
         Left            =   8100
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
         Caption         =   "datPP"
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
         Left            =   7200
         Top             =   60
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
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmPayCodeMaster.frx":0000
      Height          =   1755
      Left            =   240
      OleObjectBlob   =   "frmPayCodeMaster.frx":0014
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "M_CODE"
      Height          =   285
      Index           =   0
      Left            =   1890
      TabIndex        =   15
      Tag             =   "01-Enter Code for Attendance"
      Top             =   3930
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "ADRE"
      MaxLength       =   7
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Frequncy"
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
      Height          =   255
      Left            =   390
      TabIndex        =   19
      Top             =   4710
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Height          =   195
      Index           =   0
      Left            =   330
      TabIndex        =   18
      Top             =   3600
      Width           =   435
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "INFO:HR Code"
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
      Left            =   330
      TabIndex        =   17
      Top             =   3960
      Width           =   1275
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Type ID Code"
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
      Height          =   195
      Left            =   330
      TabIndex        =   14
      Top             =   3030
      Width           =   1560
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Type Code"
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
      Height          =   255
      Left            =   330
      TabIndex        =   11
      Top             =   2730
      Width           =   1515
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Code Name"
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
      Height          =   255
      Left            =   330
      TabIndex        =   10
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Label lblPayCode 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Pay Code"
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
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2100
      Width           =   1215
   End
End
Attribute VB_Name = "frmPayCodeMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fUPMode, fGLBNew
Dim AddChg
Dim rsDATA As New ADODB.Recordset
'
'Private Sub chkUploaded_GotFocus()
'    'Hemu - 05/13/2003 Begin
'    Call SetPanHelp(ActiveControl)
'    'Hemu - 05/13/2003 End
'End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
    On Error GoTo Mod_Err
    
    Call ST_UPD_MODE(True)
    'clpPAYP.SetFocus
    AddChg = "C"
    fGLBNew = False
    Exit Sub
    
Mod_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HR_PAY_PERIOD", "Modify")
End Sub

Private Sub cmdPrint_Click()
    cmdPrint.Enabled = False
    
    Me.vbxCrystal.WindowTitle = "Pay Period Master Report"
    Me.vbxCrystal.BoundReportHeading = "Pay Period Master Report"
    
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        vbxCrystal.Connect = "PWD=petman;"
    End If
    vbxCrystal.ReportFileName = glbIHRREPORTS & "rgridpp.rpt"
    
    Me.vbxCrystal.Action = 1
    cmdPrint.Enabled = True
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

Me.Show
AddChg = " "
datPP.ConnectionString = glbAdoIHRDB
datPP.RecordSource = "SELECT * FROM HR_PAY_CODE"
datPP.Refresh
Screen.MousePointer = vbHourglass
Me.vbxTrueGrid.SetFocus

If datPP.Recordset.BOF And datPP.Recordset.EOF Then
   cmdModify.Enabled = False
Else
   cmdModify.Enabled = True
   datPP.Recordset.MoveFirst
End If

Call INI_Controls(Me)
Display_Value
Screen.MousePointer = DEFAULT

' TODO - Replace True with security check for Inquire/Maintain
If True Then
    Call ST_UPD_MODE(False)
Else
    Call ST_UPD_MODE(False)
    cmdModify.Enabled = False
    
End If

End Sub

Private Sub ST_UPD_MODE(YN)
    Dim TF As Boolean, FT As Boolean
    
    If YN Then TF = True Else TF = False
    FT = Not TF
    
    cmdOK.Enabled = TF
    cmdCancel.Enabled = TF
'    Me.txtYear.Enabled = TF
'    Me.dlpFrom.Enabled = TF
'    Me.dlpTo.Enabled = TF
'    Me.chkUploaded.Enabled = TF
'    clpPAYP.Enabled = TF
    cmdClose.Enabled = FT
    cmdModify.Enabled = FT
'    cmdNew.Enabled = FT
'    cmdDelete.Enabled = FT
    cmdPrint.Enabled = FT
    vbxTrueGrid.Enabled = FT
    
    If datPP.Recordset.EOF Or datPP.Recordset.BOF Then
        cmdModify.Enabled = False
'        cmdDelete.Enabled = False
    End If
    cmdPrint.Visible = Not glbtermopen
    fUPMode = TF    ' update mode
End Sub

Private Sub cmdNew_Click()
    Dim SQLQ As String
    
    Call ST_UPD_MODE(True)
'    clpPAYP.SetFocus
    
    On Error GoTo AddN_Err
    Call Set_Control("B", Me)
    rsDATA.AddNew
    
    AddChg = "A"
    fGLBNew = True
    Exit Sub
    
AddN_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_PAY_PERIOD", "Add")
End Sub

Private Sub cmdOK_Click()
Dim X, xID
Dim rsCOM As New ADODB.Recordset
On Error GoTo Add_Err

If Not chkEPP() Then Exit Sub

rsDATA("PP_COMPNO") = "001"
rsDATA("PP_LUSER") = glbUserID
rsDATA("PP_LTIME") = Format(Now, "Short Time")
rsDATA("PP_LDATE") = Format(Now, "Short Date")
Call Set_Control("U", Me, rsDATA)
gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans
If glbSQL Or glbOracle Then rsDATA.Requery Else rsDATA.Resync
xID = rsDATA("PP_ID")
datPP.Refresh
datPP.Recordset.Find "PP_ID=" & xID

Call ST_UPD_MODE(False)
Me.vbxTrueGrid.SetFocus
Exit Sub

Add_Err:
If Err = 3022 Then
     MsgBox "Duplicate record existed - not entered"
     Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PAY_PERIOD", "Update")
If False Then
    Resume
End If
End Sub

Private Function chkEPP() As Boolean
'    If txtYear.Text = "" Then
'        MsgBox "Year is a required field.", vbInformation + vbOKOnly, "Missing Information"
'        txtYear.SetFocus
'        Exit Function
'    End If
'    If Not IsNumeric(txtYear.Text) Then
'        MsgBox "Year must be numeric.", vbInformation + vbOKOnly, "Bad Year"
'        txtYear.SetFocus
'        Exit Function
'    End If
'    If Val(txtYear.Text) < 1950 Or Val(txtYear.Text) > 2090 Then
'        MsgBox "Year must be between 1950 and 2090.", vbInformation + vbOKOnly, "Bad Year"
'        txtYear.SetFocus
'        Exit Function
'    End If
'    If dlpFrom = "" Then
'        MsgBox "Start Date is a required field.", vbInformation + vbOKOnly, "Missing Information"
'        dlpFrom.SetFocus
'        Exit Function
'    End If
'    If Not IsDate(dlpFrom.Text) Then
'        MsgBox "Start Date is not a valid date.", vbInformation + vbOKOnly, "Missing Information"
'        dlpFrom.SetFocus
'        Exit Function
'    End If
'    If dlpTo = "" Then
'        MsgBox "End Date is a required field.", vbInformation + vbOKOnly, "Missing Information"
'        dlpTo.SetFocus
'        Exit Function
'    End If
'    If Not IsDate(dlpTo.Text) Then
'        MsgBox "End Date is not a valid date.", vbInformation + vbOKOnly, "Missing Information"
'        dlpTo.SetFocus
'        Exit Function
'    End If
'
'    'Hemu - 05/13/2003 Begin - Start Date and End Date
'    If DaysBetween(dlpFrom, dlpTo) < 0 Then
'        MsgBox "End Date cannot be prior to Start Date.", vbInformation + vbOKOnly, "Missing Information"
'        dlpTo.SetFocus
'        Exit Function
'    End If
'    'Hemu - 05/13/2003 End
'
'    chkEPP = True
End Function

Private Sub cmdCancel_Click()
    On Error GoTo Can_Err
    
    rsDATA.CancelUpdate
    Call Display_Value
    Call ST_UPD_MODE(False)  ' reset screen's attributes
    Me.vbxTrueGrid.SetFocus
    Exit Sub
    
Can_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Cancel", "HR_PAY_PERIOD", "Cancel")
End Sub

Private Sub cmdDelete_Click()
    If datPP.Recordset.BOF And datPP.Recordset.EOF Then
        MsgBox "Nothing to Delete"
        Exit Sub
    End If
    
    On Error GoTo Del_Err
    If MsgBox("Are you sure you want to delete the selected record?", vbOKCancel, "Confirm Delete") <> vbOK Then Exit Sub
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    datPP.Refresh
    If datPP.Recordset.EOF And datPP.Recordset.BOF Then
        Call Display_Value
    End If
    Call ST_UPD_MODE(False)
    Exit Sub
    
Del_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HR_PAY_PERIOD", "Delete")
End Sub

Private Sub Display_Value()
    Dim SQLQ
    If datPP.Recordset.EOF Or datPP.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open datPP.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_PAY_CODE"
    SQLQ = SQLQ & " where PAY_CODE = " & datPP.Recordset!PAY_CODE
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub


Private Sub txtYear_GotFocus()
    'Hemu - 05/13/2003 Begin
    Call SetPanHelp(Me.ActiveControl)
    'Hemu - 05/13/2003 End
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub
