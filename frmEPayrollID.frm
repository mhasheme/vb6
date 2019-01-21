VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEPayrollID 
   Caption         =   "Payroll ID Lookup"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10020
   LinkTopic       =   "Form1"
   ScaleHeight     =   6255
   ScaleWidth      =   10020
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   9
      Tag             =   "00-Search for Surname"
      Top             =   4740
      Width           =   1935
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "Sort by Employee Number"
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
      Left            =   5400
      TabIndex        =   8
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   4680
      Width           =   2415
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
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
      Left            =   4320
      TabIndex        =   7
      Tag             =   "Find Employee"
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdPIDSort 
      Appearance      =   0  'Flat
      Caption         =   "Sort by Payroll ID"
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
      Left            =   8040
      TabIndex        =   1
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdSelHeader 
      Caption         =   "Sort By Selected Header"
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
      Left            =   5400
      TabIndex        =   0
      Top             =   5160
      Width           =   2415
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmEPayrollID.frx":0000
      Height          =   4455
      Left            =   0
      OleObjectBlob   =   "frmEPayrollID.frx":0014
      TabIndex        =   2
      Top             =   0
      Width           =   9975
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   5595
      Width           =   10020
      _Version        =   65536
      _ExtentX        =   17674
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
      Begin VB.CommandButton cmdOK 
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
         Left            =   120
         TabIndex        =   6
         Tag             =   "Select the Employee listed above"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
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
         Left            =   960
         TabIndex        =   5
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   825
      End
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
         Left            =   1920
         TabIndex        =   4
         Tag             =   "Print the Employee Listing"
         Top             =   150
         Width           =   735
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   4560
         Top             =   240
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6495
         Top             =   150
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
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.Label lblSearchBy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Surname"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   4800
      Width           =   1665
   End
End
Attribute VB_Name = "frmEPayrollID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EEList_Snap As New ADODB.Recordset
Dim EESNameSort As Integer
Dim fMarks As New Collection
Dim varMultiSelect As Boolean
Dim varHideHired
Dim EEPIDSort As Boolean
Dim SortByFieldName, xFlagSortBySel As Boolean

Private Sub cmdCancel_Click()

glbEEOK = False
'glbLEE_FName = "None Selected"
'glbLEE_SName = "None Selected"
'MsgBox Str(glbLEE_ID)
'Unload Me
glbUserUploadMode = UploadFormWithoutCheck: Unload Me
End Sub

Private Sub cmdEESort_Click()
Dim xStr
txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

xFlagSortBySel = False
EEPIDSort = False

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort.Caption = "Sort by Surname "
    glbSort = "NUMBER"
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search by Surname"
    cmdEESort.Caption = "Sort by Employee Number"
    glbSort = "NAME"
End If

If EEList() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = lblSearchBy.Caption '"Search by Surname "    'laura jan 05 1998

End Sub

Private Sub cmdEESort_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim Sch As String, SQLQ As String
Dim bkmark
Dim rsEmpNum As New ADODB.Recordset, xStr, xSelect

On Error GoTo Srch_Err
Call get_Marks
Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch, "'", "''")
    If Not xFlagSortBySel Then
        If EEPIDSort Then
            SQLQ = "ED_PAYROLL_ID  >= '" & Sch & "'"
        Else
            If EESNameSort = True Then
                If glbOracle Or glbSQL Then
                    SQLQ = "UPPER(ED_SURNAME)  >= '" & UCase(Sch) & "'"
                Else
                    SQLQ = "ED_SURNAME  >= '" & Sch & "'"
                End If
                'Find this Employee# and then to search it. In order to find the "O'R" as Surname problem #7994
                xStr = "SELECT ED_EMPNBR FROM HREMP WHERE " & SQLQ
                xStr = xStr & " AND " & glbSeleDeptUn
                If glbOracle Then
                    xStr = xStr & " ORDER BY UPPER(ED_SURNAME), UPPER(ED_FNAME)"
                Else
                    xStr = xStr & " ORDER BY ED_SURNAME, ED_FNAME"
                End If
                rsEmpNum.Open xStr, gdbAdoIhr001, adOpenStatic
                If Not rsEmpNum.EOF Then
                    xSelect = "ED_EMPNBR=" & rsEmpNum("ED_EMPNBR")
                Else
                    xSelect = " (1=1)"
                End If
                rsEmpNum.Close
                SQLQ = xSelect
                'Find this Employee# - End
            Else
                If Not glbLinamar And Not IsNumeric(txtEESearch) Then
                    Beep
                    MsgBox "Employee Identification must be numeric"
                    Exit Sub
                End If
                If glbLinamar Then
                    SQLQ = "EMPNBR >= '" & Sch & "'"
                Else
                    If glbOracle Then
                        SQLQ = "ED_EMPNBR >= '" & Sch & "'"
                    Else
                        SQLQ = "ED_EMPNBR >= " & Sch & ""
                    End If
                End If
            End If
        End If
    Else
        If SortByFieldName = "Ed_Doh" Then 'For Date Field
            If IsDate(Sch) Then
                SQLQ = SortByFieldName & " >= '" & Sch & "'"
            Else
                SQLQ = SortByFieldName & " >= '01/01/1800'"
            End If
        Else
            SQLQ = SortByFieldName & " >= '" & Sch & "'"
        End If
    End If
    Data1.Recordset.Find SQLQ
End If
If Data1.Recordset.EOF Then
    If Data1.Recordset.RecordCount > 0 Then Data1.Recordset.MoveFirst
    MsgBox "Employee not found"
End If
Call set_Marks
Screen.MousePointer = DEFAULT
cmdOK.SetFocus

Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EEList", "HREMP", "Find Next")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub cmdFind_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim X
' set the global last EE number for returned value
glbEEOK = True
If Data1.Recordset.EOF And Data1.Recordset.BOF Then 'laura 03/04/98
    Exit Sub
End If
If glbLEE_SName = "Multi_EMP" Then
    If vbxTrueGrid.SelBookmarks.count = 0 Then
        glbLEE_FName = Data1.Recordset!EmpNbr
    Else
        If vbxTrueGrid.SelBookmarks.count > 1000 Then
            MsgBox vbxTrueGrid.SelBookmarks.count & " employees are selected" + Chr(10) + " Please make that less than 1000 employees"
            Exit Sub
        End If
        glbLEE_FName = ""
        For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
            vbxTrueGrid.Bookmark = vbxTrueGrid.SelBookmarks(X)
            glbLEE_FName = glbLEE_FName & Data1.Recordset!EmpNbr & ","
        Next
        glbLEE_FName = Left(glbLEE_FName, Len(glbLEE_FName) - 1)
    End If
Else
    glbLEE_ID = Data1.Recordset("ED_EMPNBR")
    glbEmpCountry = UCase(Data1.Recordset("ED_COUNTRY"))
    
    If IsNull(Data1.Recordset("ED_ORG")) Then
        glbUNION = ""
    Else
        glbUNION = Data1.Recordset("ED_ORG")
    End If
    If Not IsNull(Data1.Recordset("ED_FNAME")) Then
        glbLEE_FName = Data1.Recordset("ED_FNAME")
    Else
        glbLEE_FName = "*ERROR*"
    End If
    If Not IsNull(Data1.Recordset("ED_SURNAME")) Then
        glbLEE_SName = Data1.Recordset("ED_SURNAME")
    Else
        glbLEE_SName = "*ERROR*"
    End If
    If glbWFC Then 'Get the glbBand
        glbBand = get_band(glbLEE_ID)
    End If
    If glbLinamar Then
        glbLEE_ProdLine = Mid(Data1.Recordset("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", Data1.Recordset("PROD_LINE")) 'Ticket #14775
    End If
End If

Unload Me

End Sub
Private Function get_band(EmpNo)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
    get_band = ""
    SQLQ = "SELECT SH_EMPNBR,SH_BAND FROM HR_SALARY_HISTORY WHERE SH_CURRENT <>0 AND SH_EMPNBR = " & EmpNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("SH_BAND")) Then
            get_band = rsTemp("SH_BAND")
        End If
    End If
    rsTemp.Close
End Function
Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPIDSort_Click()
Dim xStr
txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

xFlagSortBySel = False
EEPIDSort = True

lblSearchBy.Caption = "Search by Payroll ID"

If EEList() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = lblSearchBy.Caption '"Search by Surname "    'laura jan 05 1998

End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String
Dim dscGroup$
    
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
    
    cmdPrint.Enabled = False
    glbiOneWhere = False
    glbstrSelCri = ""
    Call glbCri_DeptUN("")
    If Len(glbstrSelCri) >= 0 Then
        Me.vbxCrystal.SelectionFormula = glbstrSelCri
    End If
    RHeading = "Employee Listing"
    Me.vbxCrystal.WindowTitle = RHeading & " Report"
    Me.vbxCrystal.BoundReportHeading = RHeading
    'Add by Frank Mar 14,2002
    If lblSearchBy.Caption = "Search by Surname" Then
        Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Else
        Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HREMP.ED_EMPNBR};ANYCHANGE;A"
    End If
    'Add by Frank Mar 14,2002
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfind.rpt"
    dscGroup$ = "PgHeading" & "= '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
    Me.vbxCrystal.Formulas(0) = dscGroup$
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
    End If
    Me.vbxCrystal.Action = 1
    cmdPrint.Enabled = True
End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function EEList()
Dim SQLQ As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

On Error GoTo EEList_Err
EEList = False
SQLQ = "Select ED_DEPTNO, "
If glbLinamar Then
    SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
    SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
ElseIf glbOracle Then
    SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
Else
    SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
End If
SQLQ = SQLQ & "ED_SURNAME, ED_FNAME,"
SQLQ = SQLQ & "ED_ALIAS, ED_BADGEID," 'New
SQLQ = SQLQ & "ED_EMPNBR, ED_PAYROLL_ID,ED_COUNTRY,"
If glbLinamar Then SQLQ = SQLQ & "SUBSTRING(ED_HOMELINE,4,16) AS ED_HOMELINE,"
If glbLinamar Then SQLQ = SQLQ & "SUBSTRING(ED_REGION,4,16) AS ED_REGION,"
SQLQ = SQLQ & "ED_INTEL, ED_EMP, ED_DOH, ED_PT, ED_ORG From HREMP"
SQLQ = SQLQ & " Where " & glbSeleDeptUn

If Not xFlagSortBySel Then
    If EEPIDSort Then 'Frank 09/24/04 Ticket #6962 Add Sort by Payroll ID function
        SQLQ = SQLQ & " ORDER BY ED_PAYROLL_ID"
    Else
        If cmdEESort.Caption = "Sort by Employee Number" Then
            If glbOracle Then
                SQLQ = SQLQ & " ORDER BY UPPER(ED_SURNAME), UPPER(ED_FNAME)"
            Else
                SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
            End If
        Else
            If glbLinamar Then
                SQLQ = SQLQ & " ORDER BY EMPNBR"
            Else
                SQLQ = SQLQ & " ORDER BY ED_EMPNBR"
            End If
        End If
    End If
Else
    SQLQ = SQLQ & " ORDER BY " & SortByFieldName
End If
Data1.RecordSource = SQLQ
Data1.Refresh
Me.vbxTrueGrid.Refresh

EEList = True

Exit Function

EEList_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EEList", "HREMP", "Select")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function



Private Sub cmdSelHeader_Click()
If vbxTrueGrid.SelStartCol <> -1 Then
    Call SortByHeader(vbxTrueGrid.SelStartCol)
Else
    MsgBox "No Header Selected"
End If
End Sub
Private Sub SortByHeader(xNo)
Dim xStr, xTitle
txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

'EEPIDSort = True

SortByFieldName = ""
Select Case xNo
Case 0
    xTitle = "Search by Surname" '"Search by Surname"
    EESNameSort = True
    'lblSearchBy.Caption = "Search by Surname"
    cmdEESort.Caption = "Sort by Employee Number"
    glbSort = "NAME"
    xFlagSortBySel = False
Case 3
    xTitle = "Search by Emp. #" '"Sort by Employee Number"
    EESNameSort = False
    'lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort.Caption = "Sort by Surname "
    glbSort = "NUMBER"
    xFlagSortBySel = False
Case 4
    xTitle = "Search by Payroll ID"
    EEPIDSort = True
    xFlagSortBySel = False
Case Else
    xTitle = "Search by " & Me.vbxTrueGrid.Columns(xNo).Caption
    SortByFieldName = Me.vbxTrueGrid.Columns(xNo).DataField
    xFlagSortBySel = True
End Select
lblSearchBy.Caption = xTitle

If EEList() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = lblSearchBy.Caption '"Search by Surname "    'laura jan 05 1998

End Sub
Private Sub Form_Activate()

If glbBasicChg Then
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
    If EEList() = False Then     ' get the info for this person
        Exit Sub
    End If          ' dpartment specific and populate the list
    MDIMain.panHelp(0).Caption = " "
    Screen.MousePointer = DEFAULT
    glbBasicChg% = False
End If
If glbCompSerial = "S/N - 2291W" Then
    vbxTrueGrid.SetFocus
    vbxTrueGrid.TabIndex = 0
Else
    txtEESearch.SetFocus
End If

End Sub

Private Sub Form_Load()
Dim xStr
'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB
'Data1.RecordSource = "HREMP"
If glbSort = "NUMBER" Then
    EESNameSort = True  'first sort is by surname
Else
    EESNameSort = False
End If

EEPIDSort = False

If glbLinamar Then
    Me.vbxTrueGrid.Columns(7).Visible = True
    Me.vbxTrueGrid.Columns(7).Caption = lStr("Region")
    Me.vbxTrueGrid.Columns(8).Caption = "Home Line"
    Me.vbxTrueGrid.Columns(8).DataField = "ED_HOMELINE"
Else
    Me.vbxTrueGrid.Columns(7).Visible = False
End If

Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Retrieving Employee List - Stand by"
Call cmdEESort_Click
If glbLEE_SName = "Multi_EMP" Then
    vbxTrueGrid.MultiSelect = 2
    If glbLEE_FName <> "" Then
        With Data1.Recordset
            If Not .EOF Then .MoveLast
            xStr = glbLEE_FName & ","
            Do Until .BOF
                If InStr(glbLEE_FName & ",", !EmpNbr & ",") <> 0 Then
                    xStr = Replace(xStr, !EmpNbr & ",", "")
                    vbxTrueGrid.SelBookmarks.Add vbxTrueGrid.Bookmark
                    DoEvents
                    If Trim(xStr) = "" Then Exit Do
                End If
                .MovePrevious
            Loop
        End With
    End If
End If
MDIMain.panHelp(0).Caption = "info:HR Main functions Locked until EE Selected"
Screen.MousePointer = DEFAULT
Me.vbxTrueGrid.Refresh

If MDIMain.lstPanel.Visible = True Then

End If

End Sub

Private Sub txtEESearch_GotFocus()
Call SetPanHelp(ActiveControl)

End Sub

Private Sub txtEESearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub vbxTrueGrid_DblClick()
frmFind = True
Call cmdOK_Click
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
        
        SQLQ = "Select ED_DEPTNO, "
        If glbLinamar Then
            SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
            SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
        ElseIf glbOracle Then
            SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
        Else
            SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
        End If
        SQLQ = SQLQ & "ED_SURNAME, ED_FNAME,"
        SQLQ = SQLQ & "ED_ALIAS, ED_BADGEID," 'New
        SQLQ = SQLQ & "ED_EMPNBR, ED_PAYROLL_ID,ED_COUNTRY,"
        If glbLinamar Then SQLQ = SQLQ & "SUBSTRING(ED_HOMELINE,4,16) AS ED_HOMELINE,"
        If glbLinamar Then SQLQ = SQLQ & "SUBSTRING(ED_REGION,4,16) AS ED_REGION,"
        SQLQ = SQLQ & "ED_INTEL, ED_EMP, ED_DOH, ED_PT, ED_ORG From HREMP"
        SQLQ = SQLQ & " Where " & glbSeleDeptUn
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    cmdOK.SetFocus
End If

End Sub

Private Sub get_Marks()
Dim X
For X = 0 To fMarks.count - 1
    fMarks.Remove 1
Next
For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
    fMarks.Add vbxTrueGrid.SelBookmarks.Item(X)
Next

End Sub
Private Sub set_Marks()
Dim X
For X = 0 To fMarks.count - 1
    vbxTrueGrid.SelBookmarks.Add fMarks.Item(X + 1)
Next
End Sub

Public Property Let HideHired(vData As APPLookupTypeEnum)
varHideHired = vData
End Property
Public Property Let MultiSelect(vData As Boolean)
varMultiSelect = vData
End Property




