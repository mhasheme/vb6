VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmBENGRLIST 
   Caption         =   "Benefit Group List"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7035
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAllEndDates 
      Appearance      =   0  'Flat
      Caption         =   "All Dates"
      Height          =   300
      Left            =   4680
      TabIndex        =   19
      Tag             =   "Save changes made"
      Top             =   5880
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdAllDates 
      Appearance      =   0  'Flat
      Caption         =   "All Dates"
      Height          =   300
      Left            =   4680
      TabIndex        =   18
      Tag             =   "Save changes made"
      Top             =   5520
      Width           =   1395
   End
   Begin VB.TextBox txtCovType 
      Appearance      =   0  'Flat
      DataField       =   "BF_COVER"
      Height          =   285
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   15
      Tag             =   "00-Type of Coverage (S, F, W, X, etc)"
      Top             =   600
      Width           =   870
   End
   Begin VB.CheckBox chkYN 
      DataField       =   "BM_CHECK"
      Height          =   285
      Left            =   2240
      TabIndex        =   1
      Top             =   5220
      Width           =   255
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fbengrpl.frx":0000
      Height          =   3945
      Left            =   120
      OleObjectBlob   =   "fbengrpl.frx":0014
      TabIndex        =   0
      Top             =   1080
      Width           =   9435
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9690
      _Version        =   65536
      _ExtentX        =   17092
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
         TabIndex        =   17
         Top             =   135
         Width           =   1305
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
         TabIndex        =   10
         Top             =   120
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1005
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
         TabIndex        =   8
         Top             =   135
         Width           =   720
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
         TabIndex        =   7
         Top             =   135
         Width           =   1245
      End
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
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   9
      Top             =   6480
      Width           =   9690
      _Version        =   65536
      _ExtentX        =   17092
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
         Left            =   1980
         TabIndex        =   4
         Tag             =   "Cancel changes made"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Update Benefit"
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
         TabIndex        =   3
         Tag             =   "Save changes made"
         Top             =   0
         Width           =   1635
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
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "BM_EDATE"
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Tag             =   "41-Effective date of salary change"
      Top             =   5520
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpProcessDate 
      Height          =   285
      Left            =   5280
      TabIndex        =   13
      Tag             =   "41-Effective date of salary change"
      Top             =   600
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpEndDate 
      Height          =   285
      Left            =   1920
      TabIndex        =   20
      Tag             =   "41-Effective date of salary change"
      Top             =   5880
      Visible         =   0   'False
      Width           =   2580
      _ExtentX        =   4551
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   5940
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Coverage Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ProcessDate"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4080
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lblACTION 
      BackStyle       =   0  'Transparent
      Caption         =   "Y/N"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   240
      TabIndex        =   12
      Top             =   5220
      Width           =   1875
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Effective Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   5580
      Width           =   1245
   End
End
Attribute VB_Name = "frmBENGRLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Option Explicit

Private Sub chkYN_KeyUp(KeyCode As Integer, Shift As Integer)
Call Update_Value
End Sub

Private Sub chkYN_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call Update_Value
End Sub

Private Sub cmdAllDates_Click() 'Ticket #18654 06/11/2010 Frank
Dim Msg As String, a%
Dim SQLQ  As String
Dim xID As Long
    If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
        If IsDate(dlpDate.Text) Then
            Msg = "Are you sure you want to update all Effective Dates with " & dlpDate.Text & " for Action " & Data1.Recordset("BM_ACTION") & "?"
            
            a% = MsgBox(Msg, 36, "Confirm Update")
            If a% <> 6 Then
                Exit Sub
            End If
            xID = Data1.Recordset("BM_BENE_ID")
            SQLQ = "UPDATE HRBENGRPLIST SET BM_EDATE = " & Date_SQL(dlpDate.Text) & " "
            SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
            SQLQ = SQLQ & "AND BM_ACTION = '" & Data1.Recordset("BM_ACTION") & "'  "
            gdbAdoIhr001.Execute SQLQ
            Data1.Refresh
            
            'Ticket #25152: Macaulay Child Development Centre - PEN Benefit only
            'If glbCompSerial = "S/N - 2420W" Then
            '    Do While Not (Data1.Recordset.EOF)
            '        If Data1.Recordset("BM_BCODE") = "PEN" Then
            '            Data1.Recordset("BF_EDATE") = CountEDate(lblEEID, Data1.Recordset("BM_WAITPERIOD"), Data1.Recordset("BM_DWM"), , , Data1.Recordset("BM_BCODE"))
            '            Data1.Recordset.Update
            '        End If
            '        Data1.Recordset.MoveNext
            '    Loop
            'End If
            
            Data1.Recordset.Find "BM_BENE_ID=" & xID
        End If
    End If
    
End Sub

Private Sub cmdAllEndDates_Click() 'Ticket #18810
Dim Msg As String, a%
Dim SQLQ  As String
Dim xID As Long
    If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
        If IsDate(dlpEndDate.Text) Then
            'Msg = "Are you sure you want to update all End Dates and Effective Dates with " & dlpEndDate.Text & " for Action EndDate?"
            Msg = "Are you sure you want to update all End Dates with " & dlpEndDate.Text & " for Action EndDate?"
            
            a% = MsgBox(Msg, 36, "Confirm Update")
            If a% <> 6 Then
                Exit Sub
            End If
            xID = Data1.Recordset("BM_BENE_ID")
            SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = " & Date_SQL(dlpEndDate.Text) & " "
            SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "' "
            SQLQ = SQLQ & "AND BM_ACTION = 'EndDate'  "
            gdbAdoIhr001.Execute SQLQ
            
            'Ticket #22214 Franks 07/11/2012 - don't update Effetive Date
            'SQLQ = "UPDATE HRBENGRPLIST SET BM_EDATE = " & Date_SQL(dlpEndDate.Text) & " "
            'SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
            'SQLQ = SQLQ & "AND BM_ACTION = 'EndDate'  "
            'gdbAdoIhr001.Execute SQLQ
                        
            Data1.Refresh
            Data1.Recordset.Find "BM_BENE_ID=" & xID
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Msg As String, a%

If glbWFC Then
    If Not chkEndDate Then
        Msg = "There are no End Dates entered for EndDate Action." & Chr(10)
        Msg = Msg & "Are you sure you want to Update Benefit?"
        a% = MsgBox(Msg, 36, "Confirm Update")
        If a% <> 6 Then
            Exit Sub
        End If
    End If
End If

Call Update_Value
If Not chkList Then Exit Sub
If IsDate(dlpProcessDate) Then
    frmEESTATS.dlpDate(4) = dlpProcessDate
End If
Call UPDBenefit
End Sub

Private Function chkEndDate()
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
Dim xID
    retVal = True
    
    SQLQ = "SELECT * FROM HRBENGRPLIST "
    SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
    SQLQ = SQLQ & "AND BM_ACTION = 'EndDate' "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xID = rsTemp("BM_BENE_ID")
        rsTemp.Close
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
        SQLQ = SQLQ & "AND BM_ACTION = 'EndDate' "
        SQLQ = SQLQ & "AND NOT (BM_ENDDATE IS NULL) "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsTemp.EOF Then
            If dlpEndDate.Visible Then 'Ticket #23948 Franks 06/20/2013
                If Not IsDate(dlpEndDate.Text) Then
                    retVal = False
                End If
            Else
                retVal = False
            End If
            Data1.Recordset.Find "BM_BENE_ID=" & xID
            Call VisibleEndDate(True)
        End If
    End If
    chkEndDate = retVal
End Function

Private Sub DateLookup1_Change()

End Sub

Private Sub dlpDate_KeyUp(KeyCode As Integer, Shift As Integer)
If IsDate(dlpDate) Then Call Update_Value
End Sub

Private Sub dlpDate_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If IsDate(dlpDate) Then Call Update_Value
End Sub

Private Sub Form_Load()
glbOnTop = "FRMBENGRLIST"
Data1.ConnectionString = glbAdoIHRDBW

If Not glbtermopen Then
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

If glbVadim Then
    lblTitle(0).Visible = True
    dlpProcessDate.Visible = True
End If

'Ticket #18810 - FOR WFC ONLY
If glbWFC Then
    vbxTrueGrid.Columns(6).Visible = True
    dlpEndDate.DataField = "BM_ENDDATE"
    Call VisibleEndDate(True)
Else
    vbxTrueGrid.Columns(6).Visible = False
    vbxTrueGrid.Columns(5).Width = 3000
End If
'Ticket #18810 - end

If Not EERetrieve() Then
    MsgBox "Sorry, Employee can not be found"
    Exit Sub
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

End Sub

Private Sub VisibleEndDate(xFlag As Boolean)
    lblTitle(2).Visible = xFlag
    dlpEndDate.Visible = xFlag
    cmdAllEndDates.Visible = xFlag
End Sub

Private Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

SQLQ = "SELECT * FROM HRBENGRPLIST "
SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
SQLQ = SQLQ & "ORDER BY BM_BCODE, BM_ACTION " 'Ticket #24178 Franks 08/02/2013
Data1.RecordSource = SQLQ
Data1.Refresh
If Data1.Recordset.EOF Then cmdOK.Enabled = False
EERetrieve = True

Exit Function
EERError:
End Function


Private Sub lblEEID_Change()
Caption = "Benefit Group List"
frmBENGRLIST.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub


Private Function chkList()
Dim SQLQ As String
Dim rsTmp As New ADODB.Recordset, xFlagDup As Boolean, xCodeT, Msg, a%

chkList = True
With Data1.Recordset
    Do Until .EOF
        chkYN = IIf(Data1.Recordset(chkYN.DataField), 1, 0)
        If IsNull(Data1.Recordset(dlpDate.DataField)) Then
            dlpDate.Text = ""
        Else
            dlpDate.Text = Data1.Recordset(dlpDate.DataField)
        End If
        If glbWFC Then 'Ticket #18810
            If IsNull(Data1.Recordset(dlpEndDate.DataField)) Then
                dlpEndDate.Text = ""
            Else
                dlpEndDate.Text = Data1.Recordset(dlpEndDate.DataField)
            End If
        End If
        If !BM_CHECK <> 0 Then
            If Not IsDate(!BM_EDATE) Then
                MsgBox "Effective date is required field."
                dlpDate.SetFocus
                chkList = False
                Exit Do
            End If
        End If
        .MoveNext
    Loop
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
    If glbVadim Then
        MsgBox "Duplicate Benefit Code entered. "
        dlpDate.SetFocus
        chkList = False
    Else
        Msg = "Duplicate Benefit Code entered. Continue? Yes/No "
        a% = MsgBox(Msg, 36, "Confirm")
        If a% <> 6 Then
            dlpDate.SetFocus
            chkList = False
        End If
    End If
End If
'Frank 11/04/2003 check duplicate benefit codes - End

End Function

Private Sub UPDBenefit()
Dim SQLQ, xACT
Dim rsBN As New ADODB.Recordset
Dim xTemp
Dim xPER
Dim xDateAge65
Dim xBenType As String

Data1.Recordset.MoveFirst
With Data1.Recordset
    Do Until .EOF
        If !BM_CHECK <> 0 Then
            xACT = Left(!BM_ACTION, 1)
            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR=" & lblEEID
            SQLQ = SQLQ & " AND BF_BCODE='" & !BM_BCODE & "'"
            If xACT <> "A" Then SQLQ = SQLQ & " AND BF_BENE_ID=" & !BM_WRKID
            rsBN.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
            If xACT = "D" And Not rsBN.EOF Then
                rsBN("BF_LUSER") = "999999998"
                rsBN.Update
                
                'Release 8.1 - List of Benefits to be deleted
                glbBenDeleted = glbBenDeleted & !BM_BCODE & ","
                
                Call AUDITBENF("D")
                rsBN.Delete
                rsBN.Close
            Else
                If xACT = "A" Then rsBN.AddNew
                
                'Release 8.1 - List of Benefits to be added
                If xACT = "A" Then
                    glbBenAdded = glbBenAdded & !BM_BCODE & ","
                Else
                    glbBenChanged = glbBenChanged & !BM_BCODE & ","
                End If
                
                rsBN("BF_COMPNO") = "001"
                rsBN("BF_EMPNBR") = lblEEID
                If glbLambton And glbVadim Then
                    rsBN("BF_PAYROLL_ID") = Get_Payroll_ID_For_Benefit(lblEEID)
                Else
                    rsBN("BF_PAYROLL_ID") = GetEmpData(lblEEID, "ED_PAYROLL_ID")
                End If
                rsBN("BF_BCODE") = !BM_BCODE
                rsBN("BF_EDATE") = !BM_EDATE
                If glbWFC And (!BM_BCODE = "HCSA4" Or !BM_BCODE = "HCSA8") Then
                    'Ticket #22108 Franks 07/12/2012 get the Coverage from Previous HCSA or HCSA1
                    xTemp = getWFCHCSA_Cover(lblEEID)
                    If Len(xTemp) > 0 Then
                        rsBN("BF_COVER") = xTemp
                    End If
                Else
                    rsBN("BF_COVER") = !BM_COVER
                End If
                rsBN("BF_AMT") = !BM_AMT
                
                'Ticket #24537
                'Ticket #18963, Ticket #24537 - more codes
                If glbCompSerial = "S/N - 2380W" Then
                    Select Case Trim(!BM_BENEFIT_GROUP)
                    Case "GHON", "GHQC", "CAMPBELL", "CAMPBC", "GHQC113", "GHON113", "CAMPBC113"
                        rsBN("BF_PPAMT") = Val(!BM_ECOST) / 52
                    Case Else
                        If (!BM_PPAMT = "0" Or !BM_PPAMT = "" Or IsNull(!BM_PPAMT)) Then
                            rsBN("BF_PPAMT") = Val(!BM_MTHECOST) / 2
                        Else
                            rsBN("BF_PPAMT") = !BM_PPAMT
                        End If
                    End Select
                Else
                    rsBN("BF_PPAMT") = !BM_PPAMT
                End If
                
                'Ticket #25500 - Goodmans - Unit Cost/Rate from Benefits Rate table
                'If glbCompSerial = "S/N - 2290W" And (rsBN("BF_BCODE") = "LIFE" Or rsBN("BF_BCODE") = "SLIFE" Or rsBN("BF_BCODE") = "CLIFE" Or rsBN("BF_BCODE") = "OLIFE") Then
                'If glbCompSerial = "S/N - 2290W" And (rsBN("BF_BCODE") = "SLIFE" Or rsBN("BF_BCODE") = "OLIFE") Then
                'Ticket #27113 - Making option to have different types of Benefit Code setup under Benefit Rates table
                xBenType = ""
                xBenType = Get_BenefitType_BenefitRateTable(rsBN("BF_BCODE"))
                If glbCompSerial = "S/N - 2290W" Then
                    'If Left(rsBN("BF_BCODE"), 1) = "S" Then
                    '    rsBN("BF_UNITCOST") = Get_BenefitRate(rsBN("BF_EMPNBR"), rsBN("BF_BCODE"), Spouse)
                    'ElseIf Left(rsBN("BF_BCODE"), 1) = "C" Then
                    '    rsBN("BF_UNITCOST") = Get_BenefitRate(rsBN("BF_EMPNBR"), rsBN("BF_BCODE"), Children)
                    'ElseIf Left(rsBN("BF_BCODE"), 1) = "O" Then
                    '    rsBN("BF_UNITCOST") = Get_BenefitRate(rsBN("BF_EMPNBR"), rsBN("BF_BCODE"), DependentRelationship.Employee)
                    'End If
                    
                    If xBenType = "S" Then
                        rsBN("BF_UNITCOST") = Get_BenefitRate(rsBN("BF_EMPNBR"), rsBN("BF_BCODE"), Spouse)
                    ElseIf xBenType = "O" Then
                        rsBN("BF_UNITCOST") = Get_BenefitRate(rsBN("BF_EMPNBR"), rsBN("BF_BCODE"), Children)
                    ElseIf xBenType = "E" Then
                        rsBN("BF_UNITCOST") = Get_BenefitRate(rsBN("BF_EMPNBR"), rsBN("BF_BCODE"), DependentRelationship.Employee)
                    Else
                        rsBN("BF_UNITCOST") = !BM_UNITCOST
                    End If
                Else
                    rsBN("BF_UNITCOST") = !BM_UNITCOST
                End If
                
                rsBN("BF_PCE") = !BM_PCE
                rsBN("BF_PCC") = !BM_PCC
                rsBN("BF_ECOST") = !BM_ECOST
                rsBN("BF_CCOST") = !BM_CCOST
                rsBN("BF_TCOST") = !BM_TCOST
                rsBN("BF_MAXDOL") = !BM_MAXDOL
                rsBN("BF_PREMIUM") = !BM_PREMIUM
                rsBN("BF_PER") = !BM_PER
                rsBN("BF_MTHCCOST") = !BM_MTHCCOST
                rsBN("BF_MTHECOST") = !BM_MTHECOST
                rsBN("BF_TAXBEN") = !BM_TAXBEN
                rsBN("BF_SALARYDEPENDANT") = !BM_SALARYDEPENDANT
                rsBN("BF_MINIMUM") = !BM_MINIMUM
                rsBN("BF_FACTOR") = !BM_FACTOR
                rsBN("BF_ROUND") = !BM_ROUND
                rsBN("BF_MAXIMUM") = !BM_MAXIMUM
                rsBN("BF_NEXTNEAREST") = !BM_NEXTNEAREST
                rsBN("BF_TAXAMOUNT") = !BM_TAXAMOUNT
                rsBN("BF_WAITPERIOD") = !BM_WAITPERIOD
                
                rsBN("BF_DWM") = !BM_DWM
                rsBN("BF_PERORDOLL") = !BM_PERORDOLL
                
                rsBN("BF_POLICY") = !BM_POLICY
                
                'Ticket #20931 - Rate Level
                rsBN("BF_RATELEVEL") = !BM_RATELEVEL

                rsBN("BF_COMMENTS") = !BM_COMMENTS
                rsBN("BF_PTAX") = !BM_PTAX
                rsBN("BF_GROUP") = !BM_BENEFIT_GROUP
                rsBN("BF_LUSER") = "999999998"
                If IsDate(dlpProcessDate) Then
                    rsBN("BF_LDATE") = dlpProcessDate
                Else
                    If xACT = "A" Then
                        rsBN("BF_LDATE") = !BM_EDATE
                    Else
                        If glbCompSerial = "S/N - 2347W" Then   'Surrey Place
                            'The !BF_LDATE should be Effective Date if the Effective Date is future date - Surrey Place
                            If CVDate(Format(!BM_EDATE, "SHORT DATE")) >= CVDate(Format(Now, "SHORT DATE")) Then
                                rsBN("BF_LDATE") = !BM_EDATE
                            Else
                                rsBN("BF_LDATE") = Date
                            End If
                        Else
                            'The Walter Fedy Partnership - Ticket #15298
                            If glbCompSerial = "S/N - 2386W" Then
                                If CVDate(Format(!BM_EDATE, "SHORT DATE")) >= CVDate(Format(Now, "SHORT DATE")) Then
                                    rsBN("BF_LDATE") = !BM_EDATE
                                Else
                                    rsBN("BF_LDATE") = Date
                                End If
                            Else
                                rsBN("BF_LDATE") = Date
                            End If
                        End If
                    End If
                End If
                rsBN("BF_LTIME") = Time$
                If glbWFC Then 'Ticket #18810
                    If IsDate(!BM_ENDDATE) Then
                        rsBN("BF_CEASEDATE") = !BM_ENDDATE
                        Call WFC_AUDIT_MANULIFE_BENF(lblEEID, !BM_BCODE, !BM_EDATE, !BM_COVER, !BM_POLICY, !BM_ENDDATE)
                    End If
                Else
                    'Ticket #25500 - Goodmans - LTD Ends Date -> 65th Birthday - 90days -> get the last day of the month
                    'If glbCompSerial = "S/N - 2290W" And rsBN("BF_BCODE") = "LTD" Then
                    If glbCompSerial = "S/N - 2290W" And rsBN("BF_BCODE") = "LTD" And ((rsBN("BF_GROUP") <> "PARTNERS" And rsBN("BF_GROUP") <> "ART") Or IsNull(rsBN("BF_GROUP"))) Then
                        'If IsNumeric(rsBN("BF_WAITPERIOD")) Then
                            'xPER = rsBN("BF_DWM")
                            'If xPER = "W" Then xPER = "ww"
                            
                            'Get the date for Age 65 or 67 based on the Benefit Group
                            If IsNull(rsBN("BF_GROUP")) Or rsBN("BF_GROUP") = "" Then
                                xDateAge65 = DateAdd("yyyy", 67, CVDate(GetEmpData(lblEEID, "ED_DOB")))
                            Else
                                xDateAge65 = DateAdd("yyyy", 65, CVDate(GetEmpData(lblEEID, "ED_DOB")))
                            End If
                            
                            'Compute LTD End Date based on employee's 65th birthday - 90days and get the last date of month
                            'rsBN("BF_CEASEDATE") = MonthLastDate(DateAdd(xPER, 0 - Val(rsBN("BF_WAITPERIOD")), CVDate(xDateAge65)))
                            'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
                            If IsNull(rsBN("BF_GROUP")) Or rsBN("BF_GROUP") = "" Then
                                rsBN("BF_CEASEDATE") = CVDate(Format("09/30/" & Year(xDateAge65), "mm/dd/yyyy"))
                            Else
                                rsBN("BF_CEASEDATE") = MonthLastDate(DateAdd("d", 0 - 90, CVDate(xDateAge65)))
                            End If
                        'End If
                    End If
                End If
                rsBN.Update
                rsBN.Close
                If xACT = "A" Then Call AUDITBENF(xACT) Else Call AUDITBENF("M")
                rsBN.Open "SELECT BF_LUSER FROM HRBENFT WHERE BF_LUSER='999999998'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
                If Not rsBN.EOF Then
                    rsBN("BF_LUSER") = glbUserID
                    rsBN.Update
                End If
                rsBN.Close
                
            End If
            
        End If
        .MoveNext
    Loop
    
    'Release 8.1 - Trim the comma at the end
    If Len(glbBenAdded) > 0 Then glbBenAdded = Left(glbBenAdded, Len(glbBenAdded) - 1)
    If Len(glbBenChanged) > 0 Then glbBenChanged = Left(glbBenChanged, Len(glbBenChanged) - 1)
    If Len(glbBenDeleted) > 0 Then glbBenDeleted = Left(glbBenDeleted, Len(glbBenDeleted) - 1)
    
End With

If glbWFC Then 'Ticket #15818
    Call WFCCNDBeneAuditFlag(glbLEE_ID)
End If

'Release 8.1 - glbBenChanged - reset if Action is Add otherwise it is included in the Email Notification incorrectly as Updated.
If glbBenChanged = "" Then
    Call updBenefitForSalDEPN(glbLEE_ID)
    glbBenChanged = ""
Else
    Call updBenefitForSalDEPN(glbLEE_ID)
End If

'Ticket #24537
If glbCompSerial = "S/N - 2380W" Then Call CalcPP

Unload Me
End Sub

Private Function AUDITBENF(ACTX)
Dim TA As New ADODB.Recordset
Dim TB As New ADODB.Recordset
Dim xPT As String, xDiv As String, xADD As String
Dim TC As New ADODB.Recordset
Dim SQLQ As String, strFields As String

On Error GoTo AUDIT_ERR
AUDITBENF = False
If glbSQL Or glbOracle Then
    SQLQ = "INSERT INTO HRAUDIT ("
    SQLQ = SQLQ & " AU_COMPNO"
    SQLQ = SQLQ & ",AU_EMPNBR"
    SQLQ = SQLQ & ",AU_PTUPL"
    SQLQ = SQLQ & ",AU_DIVUPL"
    SQLQ = SQLQ & ",AU_NEWEMP"
    SQLQ = SQLQ & ",AU_BCODE"
    SQLQ = SQLQ & ",AU_COVER"
    SQLQ = SQLQ & ",AU_MAXDOL"
    SQLQ = SQLQ & ",AU_EDATE"
    SQLQ = SQLQ & ",AU_LDATE"
    SQLQ = SQLQ & ",AU_LUSER"
    SQLQ = SQLQ & ",AU_LTIME"
    SQLQ = SQLQ & ",AU_UPLOAD"
    SQLQ = SQLQ & ",AU_TYPE"
    SQLQ = SQLQ & ",AU_TCOST"
    SQLQ = SQLQ & ",AU_PREMIUM"
    SQLQ = SQLQ & ",AU_PCE"
    SQLQ = SQLQ & ",AU_PCC"
    SQLQ = SQLQ & ",AU_PPAMT"
    SQLQ = SQLQ & ",AU_PER"
    SQLQ = SQLQ & ",AU_BAMT"
    SQLQ = SQLQ & ",AU_UNITCOST"
    SQLQ = SQLQ & ",AU_MTHECOST"
    SQLQ = SQLQ & ",AU_MTHCCOST"
    SQLQ = SQLQ & " )"
    SQLQ = SQLQ & " SELECT"
    SQLQ = SQLQ & " '001'"
    SQLQ = SQLQ & ",BF_EMPNBR"
    SQLQ = SQLQ & ",ED_PT"
    SQLQ = SQLQ & ",ED_DIV"
    SQLQ = SQLQ & ",'N'"
    SQLQ = SQLQ & ",BF_BCODE"
    SQLQ = SQLQ & ",BF_COVER"
    SQLQ = SQLQ & ",BF_MAXDOL"
    SQLQ = SQLQ & ",BF_EDATE"
    SQLQ = SQLQ & ",BF_LDATE"
    SQLQ = SQLQ & ",'" & glbUserID & "'"
    SQLQ = SQLQ & ",BF_LTIME"
    SQLQ = SQLQ & ",'N'"
    SQLQ = SQLQ & ",'" & ACTX & "'"
    SQLQ = SQLQ & ",BF_TCOST"
    SQLQ = SQLQ & ",BF_PREMIUM"
    SQLQ = SQLQ & ",BF_PCE"
    SQLQ = SQLQ & ",BF_PCC"
    SQLQ = SQLQ & ",BF_PPAMT"
    SQLQ = SQLQ & ",BF_PER"
    SQLQ = SQLQ & ",BF_AMT"
    SQLQ = SQLQ & ",BF_UNITCOST"
    SQLQ = SQLQ & ",BF_MTHECOST"
    SQLQ = SQLQ & ",BF_MTHCCOST"
    If glbOracle Then
        SQLQ = SQLQ & " FROM HRBENFT, HREMP WHERE HRBENFT.BF_EMPNBR=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " AND BF_LUSER='999999998' "
    Else
        SQLQ = SQLQ & " FROM HRBENFT INNER JOIN HREMP ON HRBENFT.BF_EMPNBR=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE BF_LUSER='999999998' "
    End If
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    TC.Open "SELECT * FROM HRBENFT WHERE BF_LUSER='999999998'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    'TA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    'strfields added by Bryan 02/Dec/05 Ticket#9899
    strFields = "AU_PTUPL, AU_DIVUPL, AU_LOC_TABL, AU_NEWEMP, AU_BCODE, AU_COVER, AU_MAXDOL, AU_EDATE, "
    strFields = strFields & "AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, AU_PPAMT, AU_PER, AU_BAMT, AU_UNITCOST, AU_MTHCCOST, "
    strFields = strFields & "AU_MTHECOST, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,"
    strFields = strFields & "AU_EMP_TABL,AU_SUPCODE_TABL,AU_ORG_TABL,AU_PAYP_TABL,AU_BCODE_TABL,AU_TREAS_TABL,AU_DOLENT_TABL,AU_EARN_TABL"
    TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    xADD = False
    Do Until TC.EOF
        TA.AddNew
        TB.Open "SELECT ED_EMPNBR,ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR=" & TC("BF_EMPNBR"), gdbAdoIhr001, adOpenKeyset
        If Not TB.EOF Then
            TA("AU_PTUPL") = TB("ED_PT")
            TA("AU_DIVUPL") = TB("ED_DIV")
        End If
        TB.Close
        TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR": TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL": TA("AU_EARN_TABL") = "EARN"
        TA("AU_NEWEMP") = "N"
        TA("AU_BCODE") = TC("BF_BCODE")
        TA("AU_COVER") = TC("BF_COVER")
        TA("AU_MAXDOL") = TC("BF_MAXDOL")
        TA("AU_EDATE") = TC("BF_EDATE")

        TA("AU_TCOST") = TC("BF_TCOST")
        TA("AU_PREMIUM") = TC("BF_PREMIUM")
        TA("AU_PCE") = TC("BF_PCE")
        TA("AU_PCC") = TC("BF_PCC")
        TA("AU_PPAMT") = TC("BF_PPAMT")
        TA("AU_PER") = TC("BF_PER")
        TA("AU_BAMT") = TC("BF_AMT")
        TA("AU_UNITCOST") = TC("BF_UNITCOST")
        TA("AU_MTHCCOST") = TC("BF_MTHCCOST")
        TA("AU_MTHECOST") = TC("BF_MTHECOST")
        TA("AU_COMPNO") = "001"
        TA("AU_EMPNBR") = TC("BF_EMPNBR")
        TA("AU_LDATE") = Date
        TA("AU_LUSER") = glbUserID
        TA("AU_LTIME") = Time$
        TA("AU_UPLOAD") = "N"
        TA("AU_TYPE") = ACTX
        TA.Update
        TC.MoveNext
    Loop
    TC.Close
End If
AUDITBENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Private Sub txtCovType_Change()
If Data1.Recordset.RecordCount = 0 Then Exit Sub
Data1.Recordset.MoveFirst
Do Until Data1.Recordset.EOF
    If txtCovType <> "" Then
        If Data1.Recordset!BM_COVER & "" <> "" And Data1.Recordset!BM_COVER & "" <> txtCovType Then
            Data1.Recordset!BM_CHECK = 0
        End If
        If Data1.Recordset!BM_COVER & "" <> "" And Data1.Recordset!BM_COVER & "" = txtCovType Then
            Data1.Recordset!BM_CHECK = 1
        End If
        
    End If
    Data1.Recordset.MoveNext
Loop
Data1.Refresh
End Sub

Private Sub txtCovType_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Call Update_Value
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    chkYN = 0
    dlpDate.Text = ""
Else
    chkYN = IIf(Data1.Recordset(chkYN.DataField), 1, 0)
    If IsNull(Data1.Recordset(dlpDate.DataField)) Then
        dlpDate.Text = ""
    Else
        dlpDate.Text = Data1.Recordset(dlpDate.DataField)
    End If
    If glbWFC Then 'Ticket #18810
    Dim xtmFlag As Boolean
        If IsNull(Data1.Recordset(dlpEndDate.DataField)) Then
            dlpEndDate.Text = ""
        Else
            dlpEndDate.Text = Data1.Recordset(dlpEndDate.DataField)
        End If
        If Data1.Recordset("BM_ACTION") = "EndDate" Then
            xtmFlag = True
        Else
            xtmFlag = False
        End If
        Call VisibleEndDate(xtmFlag)
    End If
End If
End Sub
Private Sub Update_Value()
If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
    Data1.Recordset(chkYN.DataField) = chkYN <> 0
    If IsDate(dlpDate.Text) Then
        If Year(dlpDate.Text) > 1900 And Year(dlpDate.Text) < 2050 Then
            Data1.Recordset(dlpDate.DataField) = dlpDate.Text
        Else
            Data1.Recordset(dlpDate.DataField) = Null
        End If
    Else
        Data1.Recordset(dlpDate.DataField) = Null
    End If
    If glbWFC Then 'Ticket #18810
        If IsDate(dlpEndDate.Text) Then
            If Year(dlpEndDate.Text) > 1900 And Year(dlpEndDate.Text) < 2050 Then
                Data1.Recordset(dlpEndDate.DataField) = dlpEndDate.Text
            Else
                Data1.Recordset(dlpEndDate.DataField) = Null
            End If
        Else
            Data1.Recordset(dlpEndDate.DataField) = Null
        End If
    End If
    Data1.Recordset.Update
End If
End Sub

Private Function getWFCHCSA_Cover(xEmpNo)
Dim SQLQ As String
Dim rsTmpBe As New ADODB.Recordset
Dim retVal As String
    retVal = ""
    SQLQ = "SELECT BF_EMPNBR,BF_BCODE,BF_COVER FROM HRBENFT WHERE BF_EMPNBR= " & xEmpNo & " "
    SQLQ = SQLQ & "AND (BF_BCODE = 'HCSA' OR BF_BCODE = 'HCSA1') "
    rsTmpBe.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTmpBe.EOF Then
        If Not IsNull(rsTmpBe("BF_COVER")) Then
            retVal = rsTmpBe("BF_COVER")
        End If
    End If
    rsTmpBe.Close
    getWFCHCSA_Cover = retVal
End Function
