VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSHoliday 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Holiday Season Master"
   ClientHeight    =   9900
   ClientLeft      =   525
   ClientTop       =   1470
   ClientWidth     =   11235
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
   ScaleHeight     =   9900
   ScaleWidth      =   11235
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCopySection 
      Caption         =   "Copy Year to Section"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin INFOHR_Controls.CodeLookup clpSection 
      Bindings        =   "fsHoliday.frx":0000
      DataField       =   "HL_SECTION"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   4920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "HL_DATE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Tag             =   "41-Holiday Date"
      Top             =   3990
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      DataField       =   "HL_NAME"
      DataSource      =   "data1"
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
      Left            =   2580
      MaxLength       =   30
      TabIndex        =   3
      Tag             =   "00-Holiday Name"
      Top             =   4440
      Width           =   3855
   End
   Begin VB.ComboBox cmbYear 
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   3435
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   11
      Top             =   9240
      Width           =   11235
      _Version        =   65536
      _ExtentX        =   19817
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
      Begin VB.CommandButton cmdAllHolidayAtt 
         Caption         =   "Update Attendance with ALL the Holidays"
         Height          =   495
         Left            =   7200
         TabIndex        =   20
         Top             =   120
         Width           =   2175
      End
      Begin VB.CommandButton cmdHrAttendance 
         Caption         =   "Update Attendance with SELECTED Holiday(s)"
         Height          =   495
         Left            =   4800
         TabIndex        =   15
         Top             =   120
         Width           =   2295
      End
      Begin VB.CommandButton cmdDeleteYear 
         Caption         =   "Delete &Year"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Tag             =   "Delete the Year"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdDup 
         Caption         =   "D&uplicate for Next Year"
         Height          =   375
         Left            =   2100
         TabIndex        =   10
         Top             =   120
         Width           =   2355
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   9660
         Top             =   0
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
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
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fsHoliday.frx":000B
      Height          =   2775
      Left            =   300
      OleObjectBlob   =   "fsHoliday.frx":001F
      TabIndex        =   1
      Top             =   1020
      Width           =   9435
   End
   Begin INFOHR_Controls.CodeLookup clpNewSec 
      Bindings        =   "fsHoliday.frx":A57F
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   8520
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpProv 
      DataField       =   "HL_STATE"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Tag             =   "31-Province of Employment - Code"
      Top             =   5400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Tag             =   "EDPT-Category"
      Top             =   7080
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   503
      ShowUnassigned  =   1
      ShowDescription =   0   'False
      TABLName        =   "EDPT"
      MaxLength       =   0
      MultiSelect     =   -1  'True
   End
   Begin INFOHR_Controls.CodeLookup clpProvRes 
      DataField       =   "HL_PROV"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   2280
      TabIndex        =   6
      Tag             =   "31-Province of Residence - Code"
      Top             =   5880
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   4
   End
   Begin VB.Label prov 
      AutoSize        =   -1  'True
      Caption         =   "Province of Residence"
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
      Index           =   0
      Left            =   360
      TabIndex        =   23
      Top             =   5925
      Width           =   1620
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   360
      TabIndex        =   22
      Top             =   7125
      Width           =   630
   End
   Begin VB.Label lblSelCri 
      AutoSize        =   -1  'True
      Caption         =   "Selection Criteria to Update Attendance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   21
      Top             =   6720
      Width           =   3405
   End
   Begin VB.Label prov 
      AutoSize        =   -1  'True
      Caption         =   "Province of Employment"
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
      Index           =   6
      Left            =   360
      TabIndex        =   19
      Top             =   5445
      Width           =   1710
   End
   Begin VB.Label lblTitle 
      Caption         =   "New Section:"
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
      Index           =   1
      Left            =   360
      TabIndex        =   18
      Top             =   8520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblTitle 
      Caption         =   "Section:"
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
      Index           =   0
      Left            =   360
      TabIndex        =   16
      Top             =   4935
      Width           =   1455
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   4035
      Width           =   420
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Holiday Name"
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
      Left            =   360
      TabIndex        =   13
      Top             =   4485
      Width           =   990
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      Caption         =   "Year"
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
      Left            =   300
      TabIndex        =   12
      Top             =   480
      Width           =   330
   End
End
Attribute VB_Name = "frmSHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbEditMode%, oYear
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew

Private Sub clpProv_Change()
    'Release 8.1
    If Len(Trim(clpProv.Text)) > 0 Then
        clpProvRes.Text = ""
        clpProvRes.Enabled = False
    Else
        clpProvRes.Enabled = True
    End If
End Sub

Private Sub clpProvRes_Change()
    'Release 8.1
    If Len(Trim(clpProvRes.Text)) > 0 Then
        clpProv.Text = ""
        clpProv.Enabled = False
    Else
        clpProv.Enabled = True
    End If
End Sub

Private Sub cmbYear_Change()
'Dim xYear As Integer
'xYear = Val(cmbYear)
'Data1.RecordSource = "SELECT * FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & xYear
'Data1.Refresh
End Sub

Sub cmbYear_Click()
Dim xYear As Integer
Dim SQLQ
If cmbYear = "Show All Years" Then
    SQLQ = "SELECT * FROM HR_HOLIDAY ORDER BY "
    If glbWFC Then
        SQLQ = SQLQ & "HL_SECTION, "
    End If
    SQLQ = SQLQ & "HL_DATE "
    Data1.RecordSource = SQLQ '"SELECT * FROM HR_HOLIDAY ORDER BY HL_DATE"
    cmdDeleteYear.Enabled = False
    cmdDup.Enabled = False
Else
    xYear = Val(cmbYear)
    If glbOracle Then
        SQLQ = "SELECT * FROM HR_HOLIDAY WHERE TO_CHAR(HL_DATE,'YYYY')='" & xYear & "' ORDER BY "
        If glbWFC Then
            SQLQ = SQLQ & "HL_SECTION, "
        End If
        SQLQ = SQLQ & "HL_DATE "
        Data1.RecordSource = SQLQ '"SELECT * FROM HR_HOLIDAY WHERE TO_CHAR(HL_DATE,'YYYY')='" & xYear & "' ORDER BY HL_DATE"
    Else
        SQLQ = "SELECT * FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & xYear & " ORDER BY "
        If glbWFC Then
            SQLQ = SQLQ & "HL_SECTION, "
        End If
        SQLQ = SQLQ & "HL_DATE "
        Data1.RecordSource = SQLQ '"SELECT * FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & xYear & " ORDER BY HL_DATE"
    End If
    cmdDeleteYear.Enabled = True  'cmdModify.Enabled
    cmdDup.Enabled = True
End If
Data1.Refresh
Call ST_UPD_MODE(False)
End Sub

Sub cmdCancel_Click()

On Error GoTo Can_Err

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
'Call ST_UPD_MODE(False)  ' reset screen's attributes

fglbNew = False
SET_UP_MODE
'Me.vbxTrueGrid.SetFocus
Call cmbYear_Click
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg$, INo&, X%

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg$ = "Are You Sure You Want To Delete "
Msg$ = Msg$ & Chr(10) & "This Record?  "

a% = MsgBox(Msg$, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub
Dim xYear

Call Codes_Master_Integration("HOLIDAY", txtName.Text, clpSection.Text, True)

panControls.Enabled = False
xYear = Year(Data1.Recordset!HL_DATE)
Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Do Until Data1.Recordset.EOF
    If Year(Data1.Recordset!HL_DATE) = xYear Then Exit Do
    Data1.Recordset.MoveNext
Loop
If Data1.Recordset.EOF Then
    For X% = 0 To cmbYear.ListCount - 1
        If Val(cmbYear.List(X%)) = xYear Then
            cmbYear.RemoveItem X%
            cmbYear.ListIndex = 0
            Exit For
        End If
    Next
End If
Data1.Refresh
panControls.Enabled = True
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OCC_HEALTH_SAFETY", "Delete")
'Call RollBack   '10June99 js

End Sub

Private Sub cmdAllHolidayAtt_Click()
    Dim rsJobHistory As New ADODB.Recordset
    Dim rsCurSal As New ADODB.Recordset
    Dim rsTABL As New ADODB.Recordset
    Dim rsAttendance As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim SQLHRTABL As String
    Dim SQLJobHistory As String
    Dim SQLAttendance As String
    
    Dim dAD_DHRS As Double
    Dim dDate As Date
    Dim sAD_REASON As String, xReasonCode As String
    Dim lEmpNo As Long
    Dim Answer As String
    Dim iRecordCount As Integer, I
    Dim SQLQ As String
    Dim xKey, xID
    Dim xINDICATOR As Integer
    Dim recCount As Integer
    Dim DgDef, Title$, Msg$, Response%

    Title$ = "Update Attendance"
    DgDef = MB_YESNO + MB_ICONINFORMATION + MB_DEFBUTTON2  ' Describe dialog.

    If txtName.Text = "" And dlpDate.Text = "" Then
        Exit Sub
    End If
    
    'Ticket #25938 - Oshawa Community Health Centre
    If Not clpPT.ListChecker Then
        Exit Sub
    End If
    
    Answer = MsgBox("Are you sure you want to update Attendance with ALL the Holidays?", 36, "Update Attendance")
                
    If Answer <> 6 Then Exit Sub
                
    recCount = getRecordCount_Add
    If recCount > 0 Then
        Msg$ = Str(recCount)
        If recCount = 1 Then Msg$ = Msg$ & " Employee's Attendance Record " Else Msg$ = Msg$ & " Employees Attendance Records "
        Msg$ = Msg$ & "will be update with each of the Holiday record(s). " & vbCrLf & vbCrLf & "Do you want to proceed?"
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response = IDNO Then
            Exit Sub
        End If
    Else
        MsgBox "No Employee record found to add the Holiday Attendance record."
        Exit Sub
    End If
                
    If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        Data1.Recordset.MoveFirst
        
        Do
            dDate = CDate(dlpDate.Text)
            sAD_REASON = txtName.Text
            xReasonCode = "STAT"
            
            SQLHRTABL = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'STAT'"
            rsTABL.Open SQLHRTABL, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsTABL.EOF Then
                rsTABL.AddNew
                rsTABL("TB_COMPNO") = "001"
                rsTABL("TB_NAME") = "ADRE"
                rsTABL("TB_KEY") = "STAT"
                rsTABL("TB_DESC") = sAD_REASON
                rsTABL("TB_LDATE") = dDate
                rsTABL("TB_LTIME") = Time$
                rsTABL("TB_LUSER") = glbUserID
                
                'Hemu - Ticket #15050 - TB_INDICATOR - relabelled to "No Sick Ent." and they have
                'a special logic behind this. So from this update it should not be checked to YES.
                'The database design has default value of YES.
                If glbCompSerial = "S/N - 2388W" Then 'DNSSAB
                    rsTABL("TB_INDICATOR") = 0
                End If
                'Hemu - Ticket #15050 - End
                
                rsTABL.Update
                rsTABL.Close
                'Set RSTABL = Nothing
            Else
                'Hemu - Ticket #15050
                'Ticket #22270 - Kidslink
                If glbCompSerial = "S/N - 2388W" Or glbCompSerial = "S/N - 2430W" Then 'DNSSAB, KidsLink
                    xINDICATOR = rsTABL("TB_INDICATOR")
                End If
                rsTABL.Close
                'Hemu - Ticket #15050 - End
            End If
            
        
            If Not glbSQL And Not glbOracle Then Call Pause(0.5)
            
            SQLQ = "SELECT JH_EMPNBR,JH_DHRS,JH_WHRS,JH_JOB,JH_SHIFT,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 "
            'Ticket #23342 - If Multi Position then go by Default/Primary Position
            If glbMulti Then
                SQLQ = SQLQ & "AND JH_POSITION_CONTROL = 'YES' "
            End If
            
            If Len(clpSection.Text) > 0 Then
                SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION ='" & clpSection.Text & "')"
            End If
            'Ticket #18188
            'If glbCompSerial = "S/N - 2418W" Then
               If Len(Trim(clpProv.Text)) > 0 Then
                   SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PROVEMP ='" & Trim(clpProv.Text) & "')"
               End If
            'End If
            
            'Release 8.1
            If Len(Trim(clpProvRes.Text)) > 0 Then
                SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PROV ='" & Trim(clpProvRes.Text) & "')"
            End If
           
            'Ticket #25938 - Oshawa Community Health Centre
            If clpPT.Text <> "" Then
                SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "'))"
            End If
           
            SQLJobHistory = SQLQ '"SELECT JH_EMPNBR,JH_DHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
            rsJobHistory.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            
            MDIMain.panHelp(0).FloodType = 1
            MDIMain.panHelp(1).Caption = " Please Wait.....updating " & sAD_REASON
            MDIMain.panHelp(2).Caption = ""
            
            If Not rsJobHistory.EOF Then
                iRecordCount = rsJobHistory.RecordCount
                I = 0
            End If
            
            Do While Not rsJobHistory.EOF
                If I < iRecordCount Then
                    MDIMain.panHelp(0).FloodPercent = (I / iRecordCount) * 100
                    I = I + 1
                End If
                
                If Not IsNull(rsJobHistory!JH_DHRS) Then
                    dAD_DHRS = rsJobHistory!JH_DHRS
                Else
                    dAD_DHRS = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_DHRS", 0)
                End If
                
                lEmpNo = rsJobHistory!JH_EMPNBR
            
                SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_DOA =" & Date_SQL(dDate) & " "
                SQLQ = SQLQ & "AND AD_REASON = '" & xReasonCode & "' "
                SQLQ = SQLQ & "AND AD_EMPNBR = " & lEmpNo & " "
                If glbCompSerial = "S/N - 2430W" Then 'kidslink Ticket #21437 Franks 03/06/2012
                    SQLQ = SQLQ & "AND AD_JOB = '" & rsJobHistory("JH_JOB") & "' "
                End If
                If rsAttendance.State <> 0 Then rsAttendance.Close
                rsAttendance.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If rsAttendance.EOF Then
                    rsAttendance.AddNew
                    rsAttendance("AD_COMPNO") = "001"
                    rsAttendance!AD_EMPNBR = lEmpNo
                    rsAttendance!AD_DOA = dDate
                    rsAttendance!AD_REASON = xReasonCode '"STAT"
                End If
                    rsAttendance!AD_COMM = sAD_REASON   'Holiday Name
                    
                    'St. John's Rehab Ticket #14739
                    If glbCompSerial = "S/N - 2394W" Then
                        SQLQ = "SELECT ED_EMPNBR, ED_PT FROM HREMP WHERE ED_EMPNBR = " & lEmpNo
                        If rsTemp.State <> 0 Then rsTemp.Close
                        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                        rsAttendance!AD_HRS = 0
                        If Not IsNull(rsTemp("ED_PT")) Then
                            If rsTemp("ED_PT") = "FT" Then
                                rsAttendance!AD_HRS = 7.5
                            End If
                        End If
                    ElseIf glbCompSerial = "S/N - 2430W" Then 'kidslink Ticket #21437 Franks 03/06/2012
                        'Ticket #22270 - Set Default Hours to 7 if 0
                        If dAD_DHRS = 0 Then dAD_DHRS = 7
                    
                        'calculate the hours
                        rsAttendance!AD_HRS = getKidslinkHour(lEmpNo, rsJobHistory("JH_JOB"), dAD_DHRS, dDate)
                    ElseIf glbCompSerial = "S/N - 2487W" Then 'City of Kenora Ticket #29851 Franks 07/10/2017
                        rsAttendance!AD_HRS = getCityOfKenoraHour(lEmpNo, rsJobHistory("JH_JOB"), dAD_DHRS, dDate)
                    ElseIf glbCompSerial = "S/N - 2473W" Then 'Ticket #27331 - Adoptive Families Association of BC
                        'Calculate the STAT Hours
                        rsAttendance!AD_HRS = AdoptiveFamiliesAssociationOfBC_STAT_Calc(lEmpNo, dDate)
                    Else
                        rsAttendance!AD_HRS = dAD_DHRS
                    End If
                                
                    'Hemu - Ticket #15050 - AD_INDICATOR - relabelled to "No Sick Ent." and they have
                    'a special logic behind this. So from this update it should be checked in the HRTABL.
                    'The database design has default value of YES.
                    'Ticket #22270 - Kidslink
                    If glbCompSerial = "S/N - 2388W" Or glbCompSerial = "S/N - 2430W" Then 'DNSSAB, KidsLink
                        rsAttendance!AD_INDICATOR = xINDICATOR
                    End If
                    'Hemu - Ticket #15050 - End
                    
                    rsAttendance("AD_PAYROLL_ID") = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_PAYROLL_ID")
                    rsAttendance("AD_GLNO") = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_GLNO")
                    rsAttendance("AD_ORG") = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_ORG")
                    rsAttendance("AD_JOB") = rsJobHistory("JH_JOB")
                    rsAttendance("AD_DHRS") = IIf(IsNull(rsJobHistory("JH_DHRS")), 0, rsJobHistory("JH_DHRS")) 'Ticket #19434 replce "" with 0
                    rsAttendance("AD_WHRS") = IIf(IsNull(rsJobHistory("JH_WHRS")), 0, rsJobHistory("JH_WHRS")) 'Ticket #19434 replce "" with 0
                    rsAttendance("AD_SHIFT") = IIf(IsNull(rsJobHistory("JH_SHIFT")), "", rsJobHistory("JH_SHIFT"))
                    If Not IsNull(rsJobHistory("JH_REPTAU")) Then
                        rsAttendance("AD_SUPER") = rsJobHistory("JH_REPTAU")
                    End If
                    
                    SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & lEmpNo
                    rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                    If Not rsCurSal.EOF Then
                        If rsCurSal("SH_SALARY") > 0 Then
                            rsAttendance("AD_SALARY") = rsCurSal("SH_SALARY")
                            rsAttendance("AD_SALCD") = rsCurSal("SH_SALCD")
                        End If
                    End If
                    rsCurSal.Close
                    Set rsCurSal = Nothing
                    
                    rsAttendance!AD_LDATE = Date
                    rsAttendance!AD_LTIME = Time$
                    rsAttendance!AD_LUSER = glbUserID
                    rsAttendance.Update
                    
                    xKey = lEmpNo
                    xKey = xKey & "|" & Format(dDate, "dd-mmm-yyyy")
                    xKey = xKey & "|" & xReasonCode
                    xID = rsAttendance!AD_ATT_ID
                    Call Attendance_Master_Integration(xKey, xID)
                    
                'End If
                rsJobHistory.MoveNext
                            
            Loop
            rsJobHistory.Close
            Set rsJobHistory = Nothing
            If rsAttendance.State <> 0 Then rsAttendance.Close
        
            Data1.Recordset.MoveNext
        Loop Until Data1.Recordset.EOF
    End If
    
    Data1.Refresh
    
    MDIMain.panHelp(0).FloodPercent = 100
    MsgBox "Update completed!"
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

End Sub

Private Sub cmdCopySection_Click()
Dim rsTemp As New ADODB.Recordset
Dim a As Integer, Msg$, INo&, X%, SQLQ, xYear
On Error GoTo Del_Err
xYear = cmbYear.Text
If Not xYear = "Show All Years" Then
    If Len(clpSection.Text) = 0 Then
        MsgBox lStr("No Section to Copy"), vbInformation + vbOKOnly, lStr("New Section Missing")
        clpSection.SetFocus
        Exit Sub
    End If
    If Len(clpNewSec.Text) = 0 Then
        MsgBox lStr("New Section must be entered to Copy"), vbInformation + vbOKOnly, lStr("New Section Missing")
        clpNewSec.SetFocus
        Exit Sub
    End If
    SQLQ = "SELECT * FROM HR_HOLIDAY WHERE HL_SECTION='" & clpNewSec.Text & "' "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        MsgBox lStr(clpNewSec.Text) & " exists already.", vbInformation + vbOKOnly, lStr("")
        clpNewSec.SetFocus
        Exit Sub
    End If
    rsTemp.Close
    'If Data1.Recordset.EOF Then
    'End If
    panControls.Enabled = False
    Msg$ = "Are You Sure You Want To Duplicate "
    Msg$ = Msg$ & Chr(10) & lStr("For Section ") & clpNewSec.Text & "?  "
    
    
    a% = MsgBox(Msg, vbYesNo + vbQuestion, lStr("Copy Holidays to new Section"))
    If a = vbNo Then Exit Sub
    SQLQ = "INSERT INTO HR_HOLIDAY(HL_DATE,HL_NAME,HL_SECTION_TABL, HL_SECTION,HL_STATE) "
    If glbOracle Then
        SQLQ = SQLQ & " SELECT HL_DATE,HL_NAME, 'EDSE','" & clpNewSec.Text & "',HL_STATE "
        SQLQ = SQLQ & " FROM HR_HOLIDAY WHERE TO_CHAR(HL_DATE,'YYYY')='" & xYear & "' and HL_SECTION='" & clpSection.Text & "'"
    ElseIf glbSQL Then
        SQLQ = SQLQ & " SELECT HL_DATE,HL_NAME,'EDSE','" & clpNewSec.Text & "',HL_STATE "
        SQLQ = SQLQ & " FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & xYear & " and HL_SECTION='" & clpSection.Text & "'"
    Else
        SQLQ = SQLQ & " SELECT HL_DATE,HL_NAME, 'EDSE','" & clpNewSec.Text & "',HL_STATE "
        SQLQ = SQLQ & " FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & xYear & " and HL_SECTION='" & clpSection.Text & "'"
    End If
    gdbAdoIhr001.Execute SQLQ
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    If Len(clpNewSec.Text) > 0 Then
        Call Codes_Master_Integration("HOLIDAY", "ALLRECORDS", clpNewSec.Text)
    End If
    Data1.Refresh
    panControls.Enabled = True
Else
    MsgBox "You must select a year to copy from the dropdown list", vbInformation + vbOKOnly, "Select Year"
    cmbYear.SetFocus
End If
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OCC_HEALTH_SAFETY", "Delete")
Call RollBack   '10June99 js

End Sub

Private Sub cmdDeleteYear_Click()
Dim a As Integer, Msg$, INo&, X%
Dim recCount As Integer
Dim DgDef, Title$, Response%

On Error GoTo Del_Err

Title$ = "Confirm Delete Year"
DgDef = MB_YESNO + MB_ICONINFORMATION + MB_DEFBUTTON2  ' Describe dialog.

Msg$ = "Are You Sure You Want To Delete "
Msg$ = Msg$ & Chr(10) & "The " & cmbYear & " Year?  "

a% = MsgBox(Msg$, 36, "Confirm Delete Year")
If a% <> 6 Then Exit Sub

recCount = getRecordCount_Delete
If recCount > 0 Then
    Msg$ = Str(recCount)
    If recCount = 1 Then Msg$ = Msg$ & " Holiday Record for the year " Else Msg$ = Msg$ & " Holiday Records for the year "
    Msg$ = Msg$ & "will be Deleted. " & vbCrLf & vbCrLf & "Do you want to proceed?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response = IDNO Then
        Exit Sub
    End If
Else
    MsgBox "No Holiday record for the year found to delete."
    Exit Sub
End If

panControls.Enabled = False

If glbOracle Then
    gdbAdoIhr001.Execute "DELETE FROM HR_HOLIDAY WHERE TO_CHAR(HL_DATE,'YYYY')='" & cmbYear & "'"
Else
    gdbAdoIhr001.Execute "DELETE FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & cmbYear
End If

If Not glbSQL And Not glbOracle Then Call Pause(0.5)

cmbYear.RemoveItem cmbYear.ListIndex
cmbYear.ListIndex = 0

Data1.Refresh

panControls.Enabled = True

Call ST_UPD_MODE(False)

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_HOLIDAY", "Delete")
Call RollBack   '10June99 js

End Sub

Private Sub cmdDeleteYear_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub cmdDup_Click()
Dim a As Integer, Msg$, INo&, X%, SQLQ, xYear

On Error GoTo Del_Err

'xYear = cmbYear.List(cmbYear.ListCount - 1)
xYear = cmbYear

If Not xYear = "Show All Years" Then
    panControls.Enabled = False
    Msg$ = "Are You Sure You Want To Duplicate "
    Msg$ = Msg$ & Chr(10) & "For Next Year?  "
    
    a% = MsgBox(Msg$, 36, "Confirm Duplicate For Next Year")
    If a% <> 6 Then Exit Sub
    SQLQ = "INSERT INTO HR_HOLIDAY(HL_DATE,HL_NAME,HL_SECTION,HL_STATE) "
    If glbOracle Then
        SQLQ = SQLQ & " SELECT ADD_MONTHS(HL_DATE,12),HL_NAME,HL_SECTION,HL_STATE "
        SQLQ = SQLQ & " FROM HR_HOLIDAY WHERE TO_CHAR(HL_DATE,'YYYY')='" & xYear & "'"
    ElseIf glbSQL Then
        SQLQ = SQLQ & " SELECT DATEADD(Year,1,HL_DATE),HL_NAME,HL_SECTION,HL_STATE "
        SQLQ = SQLQ & " FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & xYear
    Else
        SQLQ = SQLQ & " SELECT DATEADD('yyyy',1,HL_DATE),HL_NAME,HL_SECTION,HL_STATE "
        SQLQ = SQLQ & " FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & xYear
    End If
    gdbAdoIhr001.Execute SQLQ
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    
    Call INIData
    cmbYear = xYear + 1
    'cmbYear.AddItem xYear + 1
    'cmbYear.ListIndex = cmbYear.ListCount - 1
    
    Data1.Refresh
    panControls.Enabled = True
End If

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OCC_HEALTH_SAFETY", "Delete")
Call RollBack   '10June99 js

End Sub

Private Sub cmdDup_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


'Sub cmdModify_Click()
'Dim SQLQ As String
'Call SET_UP_MODE
'oYEAR = Year(Data1.Recordset!HL_DATE)
''Call ST_UPD_MODE(True)
'
'On Error GoTo Edit_Err
'
''dlpDate.SetFocus
'
'Exit Sub
'Edit_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "HRJOBEVL", "Add")
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'End Sub



Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True
Call SET_UP_MODE
'Call ST_UPD_MODE(True)

On Error GoTo AddN_Err


Data1.Recordset.AddNew

dlpDate.SetFocus



Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "Holiday", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub



Sub cmdOK_Click()
Dim X%, xYear
Dim xYearShow
Dim Answer As String
Dim bmk As Variant

On Error GoTo OK_Err

If Not chkHoliday Then Exit Sub
If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
    bmk = Data1.Recordset.Bookmark
Else
    bmk = 0
End If

Data1.Recordset("HL_NAME") = txtName & ""
'Ticket #18188
'If glbCompSerial = "S/N - 2418W" Then
    Data1.Recordset("HL_STATE") = clpProv & ""
'End If
'Release 8.1
Data1.Recordset("HL_PROV") = clpProvRes & ""

Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Call Codes_Master_Integration("HOLIDAY", txtName.Text, clpSection.Text)
End If

xYear = Year(dlpDate.Text)
For X = 0 To cmbYear.ListCount - 1
    If xYear <= Val(cmbYear.List(X)) Then Exit For
Next
If X = cmbYear.ListCount Then
    cmbYear.AddItem xYear
Else
    If xYear < Val(cmbYear.List(X)) Then
        cmbYear.AddItem xYear, X
    End If
End If
Data1.Refresh

If oYear <> xYear Then
    Do Until Data1.Recordset.EOF
        If Year(Data1.Recordset!HL_DATE) = oYear Then Exit Do
        Data1.Recordset.MoveNext
    Loop
    If Data1.Recordset.EOF Then
        For X% = 0 To cmbYear.ListCount - 1
            If Val(cmbYear.List(X%)) = oYear And Not Val(cmbYear.List(X%)) = 0 Then
                cmbYear.RemoveItem X%
                cmbYear.ListIndex = 0
                Exit For
            End If
        Next
    End If
    Data1.Refresh
    'cmdHrAttendance.Enabled = True
End If
If bmk <> 0 Then
    Data1.Recordset.Bookmark = bmk
End If
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)

fglbEditMode% = False

Me.vbxTrueGrid.SetFocus
'Call cmbYear_Click


Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBEVL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
'Answer = MsgBox("Do you want to update attendance?", vbYesNo, "Update Attendance")
'If Answer = vbYes Then
'   cmdHrAttendence.Enabled = True
'End If
   
End Sub




Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowTitle = "Holiday Season Report"

Me.vbxCrystal.Connect = RptODBC_SQL

If cmbYear <> "Show All Years" Then
    Me.vbxCrystal.SelectionFormula = "{@xYear}=" & cmbYear
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGHOLIDY.rpt"

Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, X%

'cmdPrint.Enabled = False

Me.vbxCrystal.Reset

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Holiday Season Report"

Me.vbxCrystal.Connect = RptODBC_SQL

If cmbYear <> "Show All Years" Then
    Me.vbxCrystal.SelectionFormula = "{@xYear}=" & cmbYear
End If
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGHOLIDY.rpt"

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub


Private Sub cmdHrAttendence_Click()


End Sub

Private Sub Command1_Click()
Dim sSQL As String
glbCodeRef = True   'global call for code refresh

Call ST_UPD_MODE(True)
Data1.Refresh

sSQL = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'STAT' "

Data1.Recordset.Open sSQL, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Data1.Recordset.EOF Then
    Data1.Recordset.AddNew
    If dlpDate.Text <> "" Then
        Data1.Recordset!TB_LDATE = CDate(dlpDate.Text)
    ElseIf txtName.Text <> "" Then
            Data1.Recordset!TB_DESC = txtName.Text
    Data1.Recordset!TB_NAME = "ANDRE"
    Data1.Recordset!TB_KEY = "STAT"
    End If
Else
    If dlpDate.Text <> "" And txtName.Text <> "" Then
        Data1.Recordset!TB_LDATE = CDate(dlpDate.Text)
        Data1.Recordset!TB_DESC = txtName.Text
    End If
    Data1.Recordset.Update
End If
End Sub

Private Sub cmdHrAttendance_Click()
    'Zahoor Butt 01/05/2006 Update Attendance Procedure added

    Dim rsJobHistory As New ADODB.Recordset
    Dim rsCurSal As New ADODB.Recordset
    Dim rsTABL As New ADODB.Recordset
    Dim rsAttendance As New ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    
    Dim SQLHRTABL As String
    Dim SQLJobHistory As String
    Dim SQLAttendance As String
    
    Dim dAD_DHRS As Double
    Dim dDate As Date
    Dim sAD_REASON As String, xReasonCode As String
    Dim lEmpNo As Long
    Dim Answer As String
    Dim iRecordCount As Integer, I
    Dim SQLQ As String
    Dim xKey, xID
    Dim xINDICATOR As Integer
    Dim X As Integer
    Dim recCount As Integer
    Dim DgDef, Title$, Msg$, Response%

    Title$ = "Update Attendance"
    DgDef = MB_YESNO + MB_ICONINFORMATION + MB_DEFBUTTON2   ' Describe dialog.

    If txtName.Text = "" And dlpDate.Text = "" Then
        Exit Sub
    End If
    
    'Ticket #25938 - Oshawa Community Health Centre
    If Not clpPT.ListChecker Then
        Exit Sub
    End If
    
    Answer = MsgBox("Are you sure you want to update Attendance with the SELECTED Holiday(s)?", 36, "Update Attendance")
                
    If Answer <> 6 Then Exit Sub
                
    recCount = getRecordCount_Add
    If recCount > 0 Then
        Msg$ = Str(recCount)
        If recCount = 1 Then Msg$ = Msg$ & " Employee's Attendance Record " Else Msg$ = Msg$ & " Employees Attendance Records "
        Msg$ = Msg$ & "will be update with each of the SELECTED Holiday record(s). " & vbCrLf & vbCrLf & "Do you want to proceed?"
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response = IDNO Then
            Exit Sub
        End If
    Else
        MsgBox "No Employee record found to add the Holiday Attendance record."
        Exit Sub
    End If
                
    If vbxTrueGrid.SelBookmarks.count = 0 Then vbxTrueGrid.SelBookmarks.Add Data1.Recordset.Bookmark
    For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
        Data1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(X)
        
        dDate = CDate(dlpDate.Text)
        sAD_REASON = txtName.Text
        xReasonCode = "STAT"
        
        SQLHRTABL = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_KEY = 'STAT'"
        rsTABL.Open SQLHRTABL, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsTABL.EOF Then
            rsTABL.AddNew
            rsTABL("TB_COMPNO") = "001"
            rsTABL("TB_NAME") = "ADRE"
            rsTABL("TB_KEY") = "STAT"
            rsTABL("TB_DESC") = sAD_REASON
            rsTABL("TB_LDATE") = dDate
            rsTABL("TB_LTIME") = Time$
            rsTABL("TB_LUSER") = glbUserID
            
            'Hemu - Ticket #15050 - TB_INDICATOR - relabelled to "No Sick Ent." and they have
            'a special logic behind this. So from this update it should not be checked to YES.
            'The database design has default value of YES.
            If glbCompSerial = "S/N - 2388W" Then 'DNSSAB
                rsTABL("TB_INDICATOR") = 0
            End If
            'Hemu - Ticket #15050 - End
            
            rsTABL.Update
            rsTABL.Close
            'Set RSTABL = Nothing
        Else
            'Hemu - Ticket #15050
            'Ticket #22270 - Kidslink
            If glbCompSerial = "S/N - 2388W" Or glbCompSerial = "S/N - 2430W" Then 'DNSSAB, KidsLink
                xINDICATOR = rsTABL("TB_INDICATOR")
            End If
            rsTABL.Close
            'Hemu - Ticket #15050 - End
        End If
        
    
        If Not glbSQL And Not glbOracle Then Call Pause(0.5)
        
        SQLQ = "SELECT JH_EMPNBR,JH_DHRS,JH_WHRS,JH_JOB,JH_SHIFT,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 "
        'Ticket #23342 - If Multi Position then go by Default/Primary Position
        If glbMulti Then
            SQLQ = SQLQ & "AND JH_POSITION_CONTROL = 'YES' "
        End If
        
        If Len(clpSection.Text) > 0 Then
            SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION ='" & clpSection.Text & "')"
        End If
        'Ticket #18188
        'If glbCompSerial = "S/N - 2418W" Then
           If Len(Trim(clpProv.Text)) > 0 Then
               SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PROVEMP ='" & Trim(clpProv.Text) & "')"
           End If
        'End If
               
        'Release 8.1
        If Len(Trim(clpProvRes.Text)) > 0 Then
            SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PROV ='" & Trim(clpProvRes.Text) & "')"
        End If
       
        'Ticket #25938 - Oshawa Community Health Centre
        If clpPT.Text <> "" Then
            SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "'))"
        End If
       
        'Ticket #29377 - Skylark Children, Youth & Families - Skip employees with Salary Distribution = "Y"
        If glbCompSerial = "S/N - 2409W" Then
            SQLQ = SQLQ & " AND JH_EMPNBR NOT IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SALDIST = 'Y')"
        End If
       
       
        SQLJobHistory = SQLQ '"SELECT JH_EMPNBR,JH_DHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
        rsJobHistory.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        MDIMain.panHelp(0).FloodType = 1
        'MDIMain.panHelp(1).Caption = " Please Wait"
        MDIMain.panHelp(1).Caption = " Please Wait.....updating " & sAD_REASON
        MDIMain.panHelp(2).Caption = ""
        
        If Not rsJobHistory.EOF Then
            iRecordCount = rsJobHistory.RecordCount
            I = 0
        End If
        
        Do While Not rsJobHistory.EOF
            If I < iRecordCount Then
                MDIMain.panHelp(0).FloodPercent = (I / iRecordCount) * 100
                I = I + 1
            End If
            
            If Not IsNull(rsJobHistory!JH_DHRS) Then
                dAD_DHRS = rsJobHistory!JH_DHRS
            Else
                dAD_DHRS = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_DHRS", 0)
            End If
            
            lEmpNo = rsJobHistory!JH_EMPNBR
        
            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_DOA =" & Date_SQL(dDate) & " "
            SQLQ = SQLQ & "AND AD_REASON = '" & xReasonCode & "' "
            SQLQ = SQLQ & "AND AD_EMPNBR = " & lEmpNo & " "
            If glbCompSerial = "S/N - 2430W" Then 'kidslink Ticket #21437 Franks 03/06/2012
                SQLQ = SQLQ & "AND AD_JOB = '" & rsJobHistory("JH_JOB") & "' "
            End If
            If rsAttendance.State <> 0 Then rsAttendance.Close
            rsAttendance.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsAttendance.EOF Then
                rsAttendance.AddNew
                rsAttendance("AD_COMPNO") = "001"
                rsAttendance!AD_EMPNBR = lEmpNo
                rsAttendance!AD_DOA = dDate
                rsAttendance!AD_REASON = xReasonCode '"STAT"
            End If
                rsAttendance!AD_COMM = sAD_REASON   'Holiday Name
            
                'St. John's Rehab Ticket #14739
                If glbCompSerial = "S/N - 2394W" Then
                    SQLQ = "SELECT ED_EMPNBR, ED_PT FROM HREMP WHERE ED_EMPNBR = " & lEmpNo
                    If rsTemp.State <> 0 Then rsTemp.Close
                    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    rsAttendance!AD_HRS = 0
                    If Not IsNull(rsTemp("ED_PT")) Then
                        If rsTemp("ED_PT") = "FT" Then
                            rsAttendance!AD_HRS = 7.5
                        End If
                    End If
                    rsTemp.Close
                    Set rsTemp = Nothing
                ElseIf glbCompSerial = "S/N - 2430W" Then 'kidslink Ticket #21437 Franks 03/06/2012
                    'Ticket #22270 - Set Default Hours to 7 if 0
                    If dAD_DHRS = 0 Then dAD_DHRS = 7
                    
                    'calculate the hours
                    rsAttendance!AD_HRS = getKidslinkHour(lEmpNo, rsJobHistory("JH_JOB"), dAD_DHRS, dDate)
                ElseIf glbCompSerial = "S/N - 2487W" Then 'City of Kenora Ticket #29851 Franks 07/10/2017
                    rsAttendance!AD_HRS = getCityOfKenoraHour(lEmpNo, rsJobHistory("JH_JOB"), dAD_DHRS, dDate)
                ElseIf glbCompSerial = "S/N - 2473W" Then 'Ticket #27331 - Adoptive Families Association of BC
                    'Calculate the STAT Hours
                    rsAttendance!AD_HRS = AdoptiveFamiliesAssociationOfBC_STAT_Calc(lEmpNo, dDate)
                ElseIf glbCompSerial = "S/N - 2409W" Then   'Ticket #29377 - Skylark Children, Youth & Families
                    'Calculate STAT Hours for non FT employees and for FT it will be Hours/Day from Current Position
                    SQLQ = "SELECT ED_EMPNBR, ED_PT FROM HREMP WHERE ED_EMPNBR = " & lEmpNo
                    If rsTemp.State <> 0 Then rsTemp.Close
                    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    rsAttendance!AD_HRS = 0
                    If Not IsNull(rsTemp("ED_PT")) Then
                        If rsTemp("ED_PT") = "FT" Then
                            rsAttendance!AD_HRS = dAD_DHRS
                        Else
                            'Compute the STAT hours
                            rsAttendance!AD_HRS = getSkylark_STAT_Hours(lEmpNo, dDate, dAD_DHRS)
                        End If
                    End If
                    rsTemp.Close
                    Set rsTemp = Nothing
                Else
                    rsAttendance!AD_HRS = dAD_DHRS
                End If
                            
                'Hemu - Ticket #15050 - AD_INDICATOR - relabelled to "No Sick Ent." and they have
                'a special logic behind this. So from this update it should be checked in the HRTABL.
                'The database design has default value of YES.
                'Ticket #22270 - Kidslink
                If glbCompSerial = "S/N - 2388W" Or glbCompSerial = "S/N - 2430W" Then 'DNSSAB, KidsLink
                    rsAttendance!AD_INDICATOR = xINDICATOR
                End If
                'Hemu - Ticket #15050 - End
                
                rsAttendance("AD_PAYROLL_ID") = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_PAYROLL_ID")
                rsAttendance("AD_GLNO") = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_GLNO")
                rsAttendance("AD_ORG") = GetEmpData(rsJobHistory!JH_EMPNBR, "ED_ORG")
                rsAttendance("AD_JOB") = rsJobHistory("JH_JOB")
                rsAttendance("AD_DHRS") = IIf(IsNull(rsJobHistory("JH_DHRS")), 0, rsJobHistory("JH_DHRS")) 'Ticket #19434 replce "" with 0
                rsAttendance("AD_WHRS") = IIf(IsNull(rsJobHistory("JH_WHRS")), 0, rsJobHistory("JH_WHRS")) 'Ticket #19434 replce "" with 0
                rsAttendance("AD_SHIFT") = IIf(IsNull(rsJobHistory("JH_SHIFT")), "", rsJobHistory("JH_SHIFT"))
                If Not IsNull(rsJobHistory("JH_REPTAU")) Then
                    rsAttendance("AD_SUPER") = rsJobHistory("JH_REPTAU")
                End If
                
                SQLQ = "SELECT SH_EMPNBR, SH_SALARY,SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <> 0 AND SH_EMPNBR = " & lEmpNo
                rsCurSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsCurSal.EOF Then
                    If rsCurSal("SH_SALARY") > 0 Then
                        rsAttendance("AD_SALARY") = rsCurSal("SH_SALARY")
                        rsAttendance("AD_SALCD") = rsCurSal("SH_SALCD")
                    End If
                End If
                rsCurSal.Close
                Set rsCurSal = Nothing
                
                rsAttendance!AD_LDATE = Date
                rsAttendance!AD_LTIME = Time$
                rsAttendance!AD_LUSER = glbUserID
                rsAttendance.Update
                
                xKey = lEmpNo
                xKey = xKey & "|" & Format(dDate, "dd-mmm-yyyy")
                xKey = xKey & "|" & xReasonCode
                xID = rsAttendance!AD_ATT_ID
                Call Attendance_Master_Integration(xKey, xID)
                
            'End If
            rsJobHistory.MoveNext
                        
        Loop
        rsJobHistory.Close
        Set rsJobHistory = Nothing
        
        If rsAttendance.State <> 0 Then rsAttendance.Close
        
        DoEvents
    Next

    MDIMain.panHelp(0).FloodPercent = 100
    MsgBox "Update completed!"
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

    
''
''
''        '*** CHECK TO SEE IF TB_NAME ADRE AND TB_KEY STAT DOES NOT ALREADY EXIST THEN ADD THEM OTHERWISE DONT ***
''
''        'If RSTABL.EOF Then
''        '
''        '    RSTABL.AddNew
''        '    RSTABL("TB_COMPNO") = "001"
''        '    RSTABL("TB_NAME") = "ADRE"
''        '    RSTABL("TB_KEY") = "STAT"
''        '    RSTABL("TB_DESC") = sAD_REASON
''        '    RSTABL("TB_LDATE") = dDate
''        '    RSTABL("TB_LTIME") = Time$
''        '    RSTABL("TB_LUSER") = glbUserID
''        '    RSTABL.Update
''        '    RSTABL.Close
''        '    Set RSTABL = Nothing
''        'End If
''
''
''        '*** CHECK TO SEE IF IT IS CURRENT POSITION THEN GET HOURS/DAY VALUE ***
''
''            SQLJobHistory = "SELECT JH_EMPNBR,JH_DHRS FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0"
''            rsJobHistory.Open SQLJobHistory, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''
''            SQLAttendance = "Select * from HR_ATTENDANCE where AD_DOA =" & Date_SQL(dDate)
''            rsAttendance.Open SQLAttendance, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
''
''        If rsAttendance.RecordCount <> 0 Then
''            Answer = MsgBox("Are you sure you want to Update Attendance?", 36, "Update Attendance")
''                If Answer = 6 Then
''
''         '***  UPDATE ATTENDANCE TABLE IF SOMEBODY DELETE OR ADD NEW EMPLOYEE***
''                        Do While Not rsAttendance.EOF
''
''                           rsAttendance.Delete
''                           rsAttendance.Update
''                           rsAttendance.MoveNext
''                        Loop
''
''                       Do While Not rsJobHistory.EOF
''                           dAD_DHRS = rsJobHistory!JH_DHRS
''                           lEmpNo = rsJobHistory!JH_EMPNBR
''
''         '*** FINALLY ADD NEW DATA INTO ATTENDANCE TABLE  ***
''
''                            rsAttendance.AddNew
''
''                            rsAttendance!AD_DOA = dDate
''                            rsAttendance!AD_REASON = "STAT"
''                            rsAttendance!AD_HRS = dAD_DHRS
''                            rsAttendance!AD_EMPNBR = lEmpNo
''                            rsAttendance!AD_COMM = sAD_REASON
''                            rsAttendance.Update
''                            rsJobHistory.MoveNext
''
''                    Loop
''                    Exit Sub
''                End If
''
''         Else
''            Answer = MsgBox("Are you sure you want to update Attendance?", 36, "Update Attendance")
''
''                If Answer = 6 Then
''
''
''                   Do While Not rsJobHistory.EOF
''                           dAD_DHRS = rsJobHistory!JH_DHRS
''                           lEmpNo = rsJobHistory!JH_EMPNBR
''
''         '*** FINALLY ADD NEW DATA INTO ATTENDANCE TABLE  ***
''
''                            rsAttendance.AddNew
''
''                            rsAttendance!AD_DOA = dDate
''                            rsAttendance!AD_REASON = "STAT"
''                            rsAttendance!AD_HRS = dAD_DHRS
''                            rsAttendance!AD_EMPNBR = lEmpNo
''                            rsAttendance!AD_COMM = sAD_REASON
''                            rsAttendance.Update
''                            rsJobHistory.MoveNext
''
''                    Loop
''                End If
''         End If
'''End If
''
        

End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    glbFrmCaption$ = Me.Caption
    glbErrNum& = ErrorNumber
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRJOBEVL", "SELECT")
End Sub

Private Sub Form_Activate()
    Call SET_UP_MODE
    Call cmbYear_Click
    glbOnTop = "FRMSHOLIDAY"
End Sub

Private Sub Form_Load()

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
glbOnTop = "FRMSHOLIDAY"

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%, SQLQ

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT * FROM HR_HOLIDAY ORDER BY "
If glbWFC Then
    SQLQ = SQLQ & "HL_SECTION, "
End If
SQLQ = SQLQ & "HL_DATE "

Data1.RecordSource = SQLQ '"HR_HOLIDAY"
Data1.Refresh

Call INIData
Call INI_Controls(Me)

If vbxTrueGrid.Visible Then Me.vbxTrueGrid.SetFocus

Call setCaption(lblTitle(0))
Call setCaption(lblTitle(1))
Call setCaption(cmdCopySection)
Call setCaption(lblPT)

vbxTrueGrid.Columns(2).Caption = lStr("Section")

If glbWFC Then
    lblTitle(0).Visible = True
    lblTitle(1).Visible = True
    clpSection.Visible = True
    clpNewSec.Visible = True
    cmdCopySection.Visible = True
    vbxTrueGrid.Columns(2).Visible = True
    clpSection.DataField = "HL_SECTION"
    lblTitle(0).FontBold = True
Else
    'Ticket #18188
    'If glbCompSerial = "S/N - 2418W" Then
        prov(6).Visible = True
        clpProv.Visible = True
        vbxTrueGrid.Columns(3).Visible = True
        
        'Release 8.1
        clpProvRes.Visible = True
        vbxTrueGrid.Columns(4).Visible = True
    
    'Else
    '    prov(6).Visible = False
    '    clpProv.Visible = False
    '    vbxTrueGrid.Columns(3).Visible = False
    'End If
    
    vbxTrueGrid.Columns(2).Visible = True
End If

'Call SET_UP_MODE
'Call ST_UPD_MODE(False)
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
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmSCodes = Nothing
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

glbOHSEdit% = TF

fUPMode = TF    ' update mode
'cmdOK.Enabled = TF
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdCancel.Enabled = TF
'cmdClose.Enabled = FT
'cmdPrint.Enabled = FT
''txtName.Enabled = TF
''dlpDate.Enabled = TF
'vbxTrueGrid.Enabled = FT
'cmbYear.Enabled = TF  'FT
cmdDeleteYear.Enabled = FT
cmdDup.Enabled = FT
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'    cmdDelete.Enabled = False
'    cmdModify.Enabled = False
    If cmbYear = "Show All Years" Then cmdDup.Enabled = False
Else
'Me.cmdModify_Click
End If
If cmbYear = "Show All Years" Then
    cmdDeleteYear.Enabled = False
    cmdDup.Enabled = False
End If
If Not gSec_Upd_Holiday Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
    cmdDeleteYear.Enabled = False
    cmdDup.Enabled = False
End If
End Sub

Private Sub txtName_GotFocus()
Call SetPanHelp(Me.ActiveControl)
End Sub

'Private Sub txtDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtDate_GotFocus()
'    Call SetPanHelp(Me.ActiveControl)
'End Sub
'Private Sub txtDate_KeyPress(KeyAscii As Integer)
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
    
    SQLQ = "SELECT * FROM HR_HOLIDAY "
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdClose.SetFocus
'    End If
End If
End Sub

Private Sub INIData()
Dim rsTD As New ADODB.Recordset
Dim rsSR As New ADODB.Recordset
Dim xStr As String
Dim SQLQ
If glbOracle Then
    SQLQ = "SELECT DISTINCT TO_CHAR(HL_DATE,'YYYY') as SHOWYEAR FROM HR_HOLIDAY ORDER BY TO_CHAR(HL_DATE,'YYYY')"
Else
    SQLQ = "SELECT DISTINCT YEAR(HL_DATE) as SHOWYEAR FROM HR_HOLIDAY ORDER BY YEAR(HL_DATE)"
End If

rsTD.Open SQLQ, gdbAdoIhr001, adOpenStatic
cmbYear.Clear
cmbYear.AddItem "Show All Years"
Do Until rsTD.EOF
    cmbYear.AddItem rsTD("SHOWYEAR")
    rsTD.MoveNext
Loop
cmbYear.ListIndex = 0
Call cmbYear_Change
End Sub

Private Function chkHoliday()
chkHoliday = False
If Len(dlpDate.Text) < 1 Then
    MsgBox "Holiday Date must be entered"
    dlpDate.SetFocus
    Exit Function
Else
    If Not IsDate(dlpDate.Text) Then
        MsgBox "Holiday Date is not a valid date"
        dlpDate.SetFocus
        Exit Function
    End If
End If
If glbWFC Then
    If Len(clpSection) = 0 Then
        MsgBox lStr(lblTitle(0).Caption) & " must be entered"
        clpSection.SetFocus
        Exit Function
    End If
End If

If Len(clpSection) > 0 Then
    If Not clpSection.ListChecker Then Exit Function
End If

'Release 8.1
If Len(clpProv) > 0 And Len(clpProvRes) > 0 Then
    MsgBox "Both Province of 'Employment' and 'Residence' cannot be selected. Only one can be selected."
    clpProv.SetFocus
    Exit Function
End If

If Len(clpProv) > 0 Then
    If clpProv.Caption = "Unassigned" Then
        MsgBox "Invalid Province of Employment"
        clpProv.SetFocus
        Exit Function
    End If
End If

'Release 8.1
If Len(clpProvRes) > 0 Then
    If clpProvRes.Caption = "Unassigned" Then
        MsgBox "Invalid Province of Residence"
        clpProvRes.SetFocus
        Exit Function
    End If
End If

'If glbCompSerial = "S/N - 2418W" Then
' If Len(clpProv) = 0 Then
'        MsgBox lStr(prov(6).Caption) & " must be entered"
'        clpProv.SetFocus
'        Exit Function
'    End If
'End If

chkHoliday = True

End Function

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
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Holiday  'gSec_Upd_Security  - Ticket #14947
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
        txtName.Enabled = True
        dlpDate.Enabled = True
        prov(6).Enabled = True
        clpPT.Enabled = False
    ElseIf Data1.Recordset.EOF Then
        UpdateState = NoRecord
        TF = False
        txtName.Enabled = False
        dlpDate.Enabled = False
        prov(6).Enabled = False
        clpPT.Enabled = False
    Else
        UpdateState = OPENING
        TF = True
        txtName.Enabled = True
        dlpDate.Enabled = True
        prov(6).Enabled = True
        clpPT.Enabled = True
    End If
    
    Call ST_UPD_MODE(TF)
    Call set_Buttons(UpdateState)
    
    If Not UpdateRight Then TF = False

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    oYear = ""
    If Not Data1.Recordset.EOF Then
        If Not IsNull(Data1.Recordset!HL_DATE) Then
            oYear = Year(Data1.Recordset!HL_DATE)
        End If
    End If
End Sub

Private Function getRecordCount_Delete()
    Dim SQLQ As String
    Dim rsHoliday As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Delete = 0
    recCount = 0

    If glbOracle Then
        SQLQ = "SELECT COUNT(HL_ID) AS TOT_REC FROM HR_HOLIDAY WHERE TO_CHAR(HL_DATE,'YYYY')='" & cmbYear & "'"
    Else
        SQLQ = "SELECT COUNT(HL_ID) AS TOT_REC FROM HR_HOLIDAY WHERE YEAR(HL_DATE)=" & cmbYear
    End If
    rsHoliday.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsHoliday.EOF Then
        recCount = rsHoliday("TOT_REC")
    Else
        recCount = 0
    End If
    rsHoliday.Close
    Set rsHoliday = Nothing
    
    getRecordCount_Delete = recCount

End Function

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsJobHistory As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    SQLQ = "SELECT COUNT(JH_EMPNBR) AS TOT_REC FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 "
    'Ticket #23342 - If Multi Position then go by Default/Primary Position
    If glbMulti Then
        SQLQ = SQLQ & "AND JH_POSITION_CONTROL = 'YES' "
    End If
     
    If Len(clpSection.Text) > 0 Then
        SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION ='" & clpSection.Text & "')"
    End If
    'Ticket #18188
    'If glbCompSerial = "S/N - 2418W" Then
       If Len(Trim(clpProv.Text)) > 0 Then
           SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PROVEMP ='" & Trim(clpProv.Text) & "')"
       End If
    'End If
    
    'Release 8.1
    If Len(Trim(clpProvRes.Text)) > 0 Then
        SQLQ = SQLQ & "AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PROV ='" & Trim(clpProvRes.Text) & "')"
    End If
        
    'Ticket #25938 - Oshawa Community Health Centre
    If clpPT.Text <> "" Then
        SQLQ = SQLQ & " AND JH_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_PT IN ('" & Replace(clpPT.Text, ",", "','") & "'))"
    End If
    
    'Ticket #29377 - Skylark Children, Youth & Families - Skip employees with Salary Distribution = "Y"
    If glbCompSerial = "S/N - 2409W" Then
        SQLQ = SQLQ & " AND JH_EMPNBR NOT IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SALDIST = 'Y')"
    End If
    
    rsJobHistory.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsJobHistory.EOF Then
        recCount = rsJobHistory("TOT_REC")
    Else
        recCount = 0
    End If
    rsJobHistory.Close
    Set rsJobHistory = Nothing
    
    getRecordCount_Add = recCount

End Function

Private Function getCityOfKenoraHour(xEmpNo, xJH_JOB, xAD_DHRS, xDate) 'Ticket #29851 Franks 07/10/2017
'Stat holiday calculations will use the previous 2 pay periods approved Timesheet.
'Total hours by position are divided by 20 and the result equals the Stat Holiday hours
'Master/Attendance History files for all attendance records that are not marked as "Absent".
Dim rsTemp As New ADODB.Recordset
Dim rsTATT As New ADODB.Recordset
Dim SQLQ As String
Dim xFDate, xTDate
Dim retVal
    retVal = 0
    'get the date range of previous 2 pay periods
    SQLQ = "SELECT * FROM HR_PAYPERIOD WHERE PP_START <=" & Date_SQL(xDate) & " "
    SQLQ = SQLQ & "AND PP_END >=" & Date_SQL(xDate) & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xTDate = rsTemp("PP_START")
        xTDate = DateAdd("D", -1, xTDate)
        xFDate = DateAdd("D", -27, xTDate)
        
        'HR_ATTENDANCE
        SQLQ = "SELECT AD_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xTDate)
        'Ticket #22440 - Do not check for Job
        'SQLQ = SQLQ & " AND AD_JOB = '" & xJH_JOB & "' "
        SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE =0 )"
        SQLQ = SQLQ & " GROUP BY AD_EMPNBR "
        If rsTATT.State <> 0 Then rsTATT.Close
        rsTATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTATT.EOF Then
            If Not IsNull(rsTATT("SumHRS")) Then
                retVal = retVal + rsTATT("SumHRS") / 20
            End If
        End If
        rsTATT.Close
        
        'HR_ATTENDANCE_HIS
        SQLQ = "SELECT AH_EMPNBR, Sum(AH_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(xTDate)
        'Ticket #22440 - Do not check for Job
        'SQLQ = SQLQ & " AND AH_JOB = '" & xJH_JOB & "' "
        SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE =0 )"
        SQLQ = SQLQ & " GROUP BY AH_EMPNBR "
        If rsTATT.State <> 0 Then rsTATT.Close
        rsTATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTATT.EOF Then
            If Not IsNull(rsTATT("SumHRS")) Then
                retVal = retVal + rsTATT("SumHRS") / 20
            End If
        End If
        rsTATT.Close
        'c.  If the commuted total is greater than the "Hours per Day",
        'the "Hours per Day" is used as the Stat holiday hours.
        If retVal > xAD_DHRS Then
            retVal = xAD_DHRS
        End If
    End If
    getCityOfKenoraHour = retVal


End Function

Private Function getKidslinkHour(xEmpNo, xJH_JOB, xAD_DHRS, xDate) 'Ticket #21437 Franks 03/06/2012
'Stat holiday calculations will use the previous 2 pay periods approved Timesheet.
'Total hours by position are divided by 20 and the result equals the Stat Holiday hours
'Master/Attendance History files for all attendance records that are not marked as "Absent".
Dim rsTemp As New ADODB.Recordset
Dim rsTATT As New ADODB.Recordset
Dim SQLQ As String
Dim xFDate, xTDate
Dim retVal
    retVal = 0
    'get the date range of previous 2 pay periods
    SQLQ = "SELECT * FROM HR_PAYPERIOD WHERE PP_START <=" & Date_SQL(xDate) & " "
    SQLQ = SQLQ & "AND PP_END >=" & Date_SQL(xDate) & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xTDate = rsTemp("PP_START")
        xTDate = DateAdd("D", -1, xTDate)
        xFDate = DateAdd("D", -27, xTDate)
        
        'HR_ATTENDANCE
        SQLQ = "SELECT AD_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xTDate)
        'Ticket #22440 - Do not check for Job
        'SQLQ = SQLQ & " AND AD_JOB = '" & xJH_JOB & "' "
        SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE =0 )"
        SQLQ = SQLQ & " GROUP BY AD_EMPNBR "
        If rsTATT.State <> 0 Then rsTATT.Close
        rsTATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTATT.EOF Then
            If Not IsNull(rsTATT("SumHRS")) Then
                retVal = retVal + rsTATT("SumHRS") / 20
            End If
        End If
        rsTATT.Close
        
        'HR_ATTENDANCE_HIS
        SQLQ = "SELECT AH_EMPNBR, Sum(AH_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(xTDate)
        'Ticket #22440 - Do not check for Job
        'SQLQ = SQLQ & " AND AH_JOB = '" & xJH_JOB & "' "
        SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE =0 )"
        SQLQ = SQLQ & " GROUP BY AH_EMPNBR "
        If rsTATT.State <> 0 Then rsTATT.Close
        rsTATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTATT.EOF Then
            If Not IsNull(rsTATT("SumHRS")) Then
                retVal = retVal + rsTATT("SumHRS") / 20
            End If
        End If
        rsTATT.Close
        'c.  If the commuted total is greater than the "Hours per Day",
        'the "Hours per Day" is used as the Stat holiday hours.
        If retVal > xAD_DHRS Then
            retVal = xAD_DHRS
        End If
    End If
    getKidslinkHour = retVal
End Function

Private Function AdoptiveFamiliesAssociationOfBC_STAT_Calc(xEmpNbr, xSTATDate As Date)
    'Ticket #27331 - Adoptive Families Association of BC
    'Calculate the STAT Hours based on the following logic
        '- Employee should have been hired at least 30 days preceding the STAT Holiday
        '- Employee should have worked at least 15 days during the 30 days period prior to STAT Holiday
        '- Get total hours worked from Attendance and Attendance History
        '- Get total days worked from Attendance and Attendance History
        '- STAT Hours = TotalHoursWorked / TotalDaysWorked
        
    Dim rsAttend As New ADODB.Recordset
    Dim SQLQ As String
    Dim xHireDate As Date
    Dim x30Days As Date
    Dim xWorkedHrs
    Dim xWorkedDays

    AdoptiveFamiliesAssociationOfBC_STAT_Calc = 0
    xWorkedHrs = 0
    xWorkedDays = 0
    
    'Get Hire Date
    xHireDate = GetEmpData(xEmpNbr, "ED_DOH", "01/01/1900")
    If xHireDate = "01/01/1900" Then Exit Function
    
    'At least employed 30days before STAT Holiday
    If DateDiff("d", CVDate(xHireDate), CVDate(xSTATDate)) >= 30 Then
        'Worked at least 15days prior to STAT Holiday in 30days period preceding the STAT Holiday
        x30Days = CVDate(xSTATDate) - 30
                
        'Compute the # of days (at least 15) and hours worked during 30days period prior STAT Date, i.e. x30Days to xSTATDate
        'Attendance
        SQLQ = "SELECT AD_EMPNBR, SUM(AD_HRS) AS WORKEDHRS, COUNT(AD_DOA) AS WORKEDDAYS FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(x30Days)
        SQLQ = SQLQ & " AND AD_DOA <= " & Date_SQL(xSTATDate)
        'SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_INDICATOR <> 0)"
        SQLQ = SQLQ & " AND AD_INCID <> 0"
        SQLQ = SQLQ & " GROUP BY AD_EMPNBR "
        If rsAttend.State <> 0 Then rsAttend.Close
        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsAttend.EOF Then
            If Not IsNull(rsAttend("SumHRS")) Then
                xWorkedHrs = xWorkedHrs + rsAttend("WORKEDHRS")
                xWorkedDays = xWorkedDays + rsAttend("WORKEDDAYS")
            End If
        End If
        rsAttend.Close
        Set rsAttend = Nothing
        
        'Attendance History
        SQLQ = "SELECT AH_EMPNBR, SUM(AH_HRS) AS WORKEDHRS, COUNT(AH_DOA) AS WORKEDDAYS FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND AH_DOA >= " & Date_SQL(x30Days)
        SQLQ = SQLQ & " AND AH_DOA <= " & Date_SQL(xSTATDate)
        'SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_INDICATOR <> 0)"
        SQLQ = SQLQ & " AND AD_INCID <> 0"
        SQLQ = SQLQ & " GROUP BY AH_EMPNBR "
        If rsAttend.State <> 0 Then rsAttend.Close
        rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsAttend.EOF Then
            If Not IsNull(rsAttend("SumHRS")) Then
                xWorkedHrs = xWorkedHrs + rsAttend("WORKEDHRS")
                xWorkedDays = xWorkedDays + rsAttend("WORKEDDAYS")
            End If
        End If
        rsAttend.Close
        Set rsAttend = Nothing
        
        'STAT Hours allocation
        If xWorkedHrs > 0 And xWorkedDays >= 15 Then
            AdoptiveFamiliesAssociationOfBC_STAT_Calc = Round(xWorkedHrs / xWorkedDays, 2)
        End If
    End If
End Function

Private Function getSkylark_STAT_Hours(xEmpNo, xSTATDate As Date, xAD_DHRS)
Dim rsTemp As New ADODB.Recordset
Dim rsTATT As New ADODB.Recordset
Dim SQLQ As String
Dim xFDate, xTDate
Dim xTotHours
Dim xDaysWorked
Dim retVal

    'A. Total Hours = Sum the attendance hours from both Attendance and History with AD_SEN = 'Y' for the Previous 2 Pay Periods.
    'B. # of Days Worked = # of unique Dates Attendance record: 2 records of same date = 1 work day.
    'C. STAT Hours = Average Hours per Day = divide # of Days Worked into Total Hours
    'D.  If STAT Hours > "Hours per Day", then update with "Hours per Day".

    retVal = 0
    xTotHours = 0
    xDaysWorked = 0
    
    'A. Total Hours
    'Get the Date range of Previous 2 Pay Periods
    SQLQ = "SELECT * FROM HR_PAYPERIOD WHERE PP_START <=" & Date_SQL(xSTATDate) & " "
    SQLQ = SQLQ & "AND PP_END >=" & Date_SQL(xSTATDate) & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xTDate = rsTemp("PP_START")
        xTDate = DateAdd("D", -1, xTDate)
        xFDate = DateAdd("D", -27, xTDate)
        
        'HR_ATTENDANCE
        SQLQ = "SELECT AD_EMPNBR, Sum(AD_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xTDate)
        'SQLQ = SQLQ & " AND AD_JOB = '" & xJH_JOB & "' "
        'SQLQ = SQLQ & " AND AD_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE =0 )"
        SQLQ = SQLQ & " AND AD_SEN <> 0 "
        SQLQ = SQLQ & " GROUP BY AD_EMPNBR "
        If rsTATT.State <> 0 Then rsTATT.Close
        rsTATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTATT.EOF Then
            If Not IsNull(rsTATT("SumHRS")) Then
                xTotHours = xTotHours + rsTATT("SumHRS")
            End If
        End If
        rsTATT.Close
        Set rsTATT = Nothing
        
        'HR_ATTENDANCE_HIS
        SQLQ = "SELECT AH_EMPNBR, Sum(AH_HRS) AS SumHRS"
        SQLQ = SQLQ & " FROM HR_ATTENDANCE_HISTORY "
        SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(xTDate)
        'SQLQ = SQLQ & " AND AH_JOB = '" & xJH_JOB & "' "
        'SQLQ = SQLQ & " AND AH_REASON IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE =0 )"
        SQLQ = SQLQ & " AND AH_SEN <> 0 "
        SQLQ = SQLQ & " GROUP BY AH_EMPNBR "
        If rsTATT.State <> 0 Then rsTATT.Close
        rsTATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTATT.EOF Then
            If Not IsNull(rsTATT("SumHRS")) Then
                xTotHours = xTotHours + rsTATT("SumHRS")
            End If
        End If
        rsTATT.Close
        Set rsTATT = Nothing
        
        'B. # of Days Worked
        SQLQ = "SELECT COUNT(*) AS DAYS_WORKED FROM "
        SQLQ = SQLQ & " (SELECT DISTINCT AD_DOA FROM HR_ATTENDANCE A1 "
        SQLQ = SQLQ & " WHERE AD_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xTDate)
        SQLQ = SQLQ & " AND AD_SEN <> 0 "
        SQLQ = SQLQ & " UNION "
        SQLQ = SQLQ & " SELECT DISTINCT AH_DOA FROM HR_ATTENDANCE_HISTORY A2 "
        SQLQ = SQLQ & " WHERE AH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND AH_DOA >=" & Date_SQL(xFDate)
        SQLQ = SQLQ & " AND AH_DOA <=" & Date_SQL(xTDate)
        SQLQ = SQLQ & " AND AH_SEN <> 0 "
        SQLQ = SQLQ & " AND AH_DOA NOT IN (SELECT AD_DOA FROM HR_ATTENDANCE)) A1A2"
        rsTATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTATT.EOF Then
            If Not IsNull(rsTATT("DAYS_WORKED")) Then
                xDaysWorked = rsTATT("DAYS_WORKED")
            End If
        End If
        rsTATT.Close
        Set rsTATT = Nothing

        'C. STAT Hours
        'Average Hours per Day
        If xDaysWorked <> 0 Then
            retVal = xTotHours / xDaysWorked
        Else
            retVal = 0
        End If
        
        'D.  If STAT Hours > "Hours per Day", then update with "Hours per Day".
        If retVal > xAD_DHRS Then
            retVal = xAD_DHRS
        End If
        
    End If
    rsTemp.Close
    Set rsTemp = Nothing
    
    getSkylark_STAT_Hours = retVal

End Function

