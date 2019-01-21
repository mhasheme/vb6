VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGPPayCodeList 
   Caption         =   "GP Pay Code List"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   10770
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   4560
      Width           =   4335
      Begin VB.OptionButton Opt1 
         Caption         =   "Show All"
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   13
         Top             =   120
         Width           =   1335
      End
      Begin VB.OptionButton Opt1 
         Caption         =   "Show Non Common"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin VB.CheckBox chkYN 
      DataField       =   "IC_CHECK"
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Top             =   4680
      Width           =   255
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmGPPayCodeList.frx":0000
      Height          =   3945
      Left            =   120
      OleObjectBlob   =   "frmGPPayCodeList.frx":0014
      TabIndex        =   1
      Top             =   600
      Width           =   10515
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10770
      _Version        =   65536
      _ExtentX        =   18997
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
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
         Left            =   1200
         TabIndex        =   5
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         AutoSize        =   -1  'True
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
         TabIndex        =   4
         Top             =   135
         Width           =   720
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         DataSource      =   "Data1"
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   4680
         TabIndex        =   3
         Top             =   120
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   7
      Top             =   5700
      Width           =   10770
      _Version        =   65536
      _ExtentX        =   18997
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
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Update GP Pay Code"
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
         TabIndex        =   9
         Tag             =   "Save changes made"
         Top             =   0
         Width           =   2235
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
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
         TabIndex        =   8
         Tag             =   "Cancel changes made"
         Top             =   0
         Width           =   795
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   4080
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ConnectMode     =   3
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
         Caption         =   "Ado1"
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
   Begin VB.Label lblACTION 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Action"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   4680
      Width           =   1755
   End
End
Attribute VB_Name = "frmGPPayCodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
    Call Employee_GP_BenefitDeduction_Integration(glbLEE_ID, glbUserID, True)
    Unload Me
End Sub

Private Sub Form_Load()
glbOnTop = "frmGPPayCodeList"
Screen.MousePointer = DEFAULT
Data1.ConnectionString = glbAdoIHRDB
If Not glbtermopen Then
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If
If Not EERetrieve() Then
    MsgBox "Sorry, Employee can not be found"
    Exit Sub
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If
End Sub
Private Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

SQLQ = "SELECT * FROM HR_GP_INCOMECODE_MTR_WRK "
SQLQ = SQLQ & "WHERE IC_WRKEMP = '" & glbUserID & "'  "
SQLQ = SQLQ & "AND (IC_ACTION = 'Add' OR IC_ACTION = 'Delete') "
If Opt1(0).Value Then
    SQLQ = SQLQ & "AND IC_COMMON = 0 "
End If
'SQLQ = SQLQ & "ORDER BY IC_ACTION,IC_HRCODEGROUP " '
SQLQ = SQLQ & "ORDER BY IC_ACTION,IC_HRCODEGROUP,IC_CODETYPE_DESC,IC_INCOMECODEDESC "
Data1.RecordSource = SQLQ
Data1.Refresh
If Data1.Recordset.EOF Then
    cmdOK.Enabled = False
Else
    cmdOK.Enabled = True
End If
EERetrieve = True

Exit Function
EERError:
End Function


Private Sub lblEEID_Change()
frmGPPayCodeList.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
lblEENum = ShowEmpnbr(lblEEID)
End Sub


Private Function chkList()
Dim SQLQ As String
Dim rsTmp As New ADODB.Recordset, xFlagDup As Boolean, xCodeT, Msg, a%

chkList = True
With Data1.Recordset
'    Do Until .EOF
'        chkYN = IIf(Data1.Recordset(chkYN.DataField), 1, 0)
'        If IsNull(Data1.Recordset(dlpDate.DataField)) Then
'            dlpDate.Text = ""
'        Else
'            dlpDate.Text = Data1.Recordset(dlpDate.DataField)
'        End If
'        If !BM_CHECK <> 0 Then
'            If Not IsDate(!BM_EDATE) Then
'                MsgBox "Effective date is required field."
'                dlpDate.SetFocus
'                chkList = False
'                Exit Do
'            End If
'        End If
'        .MoveNext
'    Loop
End With
'Frank 11/04/2003 check duplicate benefit codes - Begin
SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_CHECK <> 0 "
SQLQ = SQLQ & "AND BM_WRKEMP = '" & glbUserID & "' "
SQLQ = SQLQ & "AND BM_ACTION <> 'Delete' "
SQLQ = SQLQ & "ORDER BY BM_BCODE "
xFlagDup = False: xCodeT = "**"
rsTmp.Open SQLQ, gdbAdoIhr001W, adOpenStatic
Do While Not rsTmp.EOF
    If rsTmp("BM_BCODE") = xCodeT Then
        xFlagDup = True
    Else
        xCodeT = rsTmp("BM_BCODE")
    End If
    rsTmp.MoveNext
Loop
rsTmp.Close
If xFlagDup Then
'    If glbVadim Then
'        MsgBox "Duplicate Benefit Code entered. "
'        dlpDate.SetFocus
'        chkList = False
'    Else
'        Msg = "Duplicate Benefit Code entered. Continue? Yes/No "
'        a% = MsgBox(Msg, 36, "Confirm")
'        If a% <> 6 Then
'            dlpDate.SetFocus
'            chkList = False
'        End If
'    End If
End If
'Frank 11/04/2003 check duplicate benefit codes - End

End Function
Private Sub UPDBenefit()
Dim SQLQ, xACT
Dim rsBN As New ADODB.Recordset

'Data1.Recordset.MoveFirst
'With Data1.Recordset
'    Do Until .EOF
'        If !BM_CHECK <> 0 Then
'            xACT = Left(!BM_ACTION, 1)
'            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR=" & lblEEID
'            SQLQ = SQLQ & " AND BF_BCODE='" & !BM_BCODE & "'"
'            If xACT <> "A" Then SQLQ = SQLQ & " AND BF_BENE_ID=" & !BM_WRKID
'            rsBN.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
'            If xACT = "D" And Not rsBN.EOF Then
'                rsBN("BF_LUSER") = "999999998"
'                rsBN.Update
'                Call AUDITBENF("D")
'                rsBN.Delete
'                rsBN.Close
'            Else
'                If xACT = "A" Then rsBN.AddNew
'                rsBN("BF_COMPNO") = "001"
'                rsBN("BF_EMPNBR") = lblEEID
'                rsBN("BF_BCODE") = !BM_BCODE
'                rsBN("BF_EDATE") = !BM_EDATE
'                rsBN("BF_COVER") = !BM_COVER
'                rsBN("BF_AMT") = !BM_AMT
'                rsBN("BF_PPAMT") = !BM_PPAMT
'                rsBN("BF_UNITCOST") = !BM_UNITCOST
'                rsBN("BF_PCE") = !BM_PCE
'                rsBN("BF_PCC") = !BM_PCC
'                rsBN("BF_ECOST") = !BM_ECOST
'                rsBN("BF_CCOST") = !BM_CCOST
'                rsBN("BF_TCOST") = !BM_TCOST
'                rsBN("BF_MAXDOL") = !BM_MAXDOL
'                rsBN("BF_PREMIUM") = !BM_PREMIUM
'                rsBN("BF_PER") = !BM_PER
'                rsBN("BF_MTHCCOST") = !BM_MTHCCOST
'                rsBN("BF_MTHECOST") = !BM_MTHECOST
'                rsBN("BF_TAXBEN") = !BM_TAXBEN
'                rsBN("BF_SALARYDEPENDANT") = !BM_SALARYDEPENDANT
'                rsBN("BF_MINIMUM") = !BM_MINIMUM
'                rsBN("BF_FACTOR") = !BM_FACTOR
'                rsBN("BF_ROUND") = !BM_ROUND
'                rsBN("BF_MAXIMUM") = !BM_MAXIMUM
'                rsBN("BF_NEXTNEAREST") = !BM_NEXTNEAREST
'                rsBN("BF_TAXAMOUNT") = !BM_TAXAMOUNT
'                rsBN("BF_WAITPERIOD") = !BM_WAITPERIOD
'
'                rsBN("BF_DWM") = !BM_DWM
'                rsBN("BF_PERORDOLL") = !BM_PERORDOLL
'
'                rsBN("BF_POLICY") = !BM_POLICY
'
'                rsBN("BF_COMMENTS") = !BM_COMMENTS
'                rsBN("BF_PTAX") = !BM_PTAX
'                rsBN("BF_GROUP") = !BM_BENEFIT_GROUP
'                rsBN("BF_LUSER") = "999999998"
'                If IsDate(dlpProcessDate) Then
'                    rsBN("BF_LDATE") = dlpProcessDate
'                Else
'                    If xACT = "A" Then
'                        rsBN("BF_LDATE") = !BM_EDATE
'                    Else
'                        If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
'                            'The !BF_LDATE should be Effective Date if the Effective Date is future date - Surrey Place
'                            If CVDate(Format(!BM_EDATE, "SHORT DATE")) >= CVDate(Format(Now, "SHORT DATE")) Then
'                                rsBN("BF_LDATE") = !BM_EDATE
'                            Else
'                                rsBN("BF_LDATE") = Date
'                            End If
'                        Else
'                            'The Walter Fedy Partnership - Ticket #15298
'                            If glbCompSerial = "S/N - 2386W" Then
'                                If CVDate(Format(!BM_EDATE, "SHORT DATE")) >= CVDate(Format(Now, "SHORT DATE")) Then
'                                    rsBN("BF_LDATE") = !BM_EDATE
'                                Else
'                                    rsBN("BF_LDATE") = Date
'                                End If
'                            Else
'                                rsBN("BF_LDATE") = Date
'                            End If
'                        End If
'                    End If
'                End If
'                rsBN("BF_LTIME") = Time$
'                rsBN.Update
'                rsBN.Close
'                If xACT = "A" Then Call AUDITBENF(xACT) Else Call AUDITBENF("M")
'                rsBN.Open "SELECT BF_LUSER FROM HRBENFT WHERE BF_LUSER='999999998'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
'                If Not rsBN.EOF Then
'                    rsBN("BF_LUSER") = glbUserID
'                    rsBN.Update
'                End If
'                rsBN.Close
'
'            End If
'
'        End If
'        .MoveNext
'    Loop
'End With
'If glbWFC Then 'Ticket #15818
'    Call WFCCNDBeneAuditFlag(glbLEE_ID)
'End If
'Call updBenefitForSalDEPN(glbLEE_ID)
Unload Me
End Sub

Private Sub Opt1_Click(Index As Integer)
    Call EERetrieve
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    chkYN = 0
Else
    chkYN = IIf(Data1.Recordset(chkYN.DataField), 1, 0)
End If
End Sub

Private Sub chkYN_KeyUp(KeyCode As Integer, Shift As Integer)
Call Update_Value
End Sub

Private Sub chkYN_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Update_Value
End Sub

Private Sub Update_Value()
If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
    Data1.Recordset(chkYN.DataField) = chkYN <> 0
    Data1.Recordset.Update
End If
End Sub
