VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmSFFind 
   Appearance      =   0  'Flat
   Caption         =   "Find Candidate"
   ClientHeight    =   6435
   ClientLeft      =   1065
   ClientTop       =   1455
   ClientWidth     =   10440
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6435
   ScaleWidth      =   10440
   Begin VB.CheckBox chkMassDele 
      Caption         =   "Mass Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   4800
      Width           =   1335
   End
   Begin VB.Frame fraMassDele 
      Height          =   975
      Left            =   3720
      TabIndex        =   12
      Top             =   4700
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdStart 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "Start Mass Delete"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Tag             =   "Close and exit this screen"
         Top             =   520
         Width           =   2265
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   3450
         TabIndex        =   15
         Tag             =   "40-Date upto and including this date forward"
         Top             =   200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   14
         Tag             =   "40-Date from and including this date forward"
         Top             =   200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1255
      End
      Begin VB.Label lblFromTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From / To Date"
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
         TabIndex        =   16
         Top             =   200
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSelHeader 
      Caption         =   "Sort By Selected Header"
      Height          =   375
      Left            =   7800
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   4320
      TabIndex        =   1
      Tag             =   "Find Employee"
      Top             =   4320
      Width           =   735
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "Sort by Employee Number"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   4320
      Width           =   2415
   End
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2280
      MaxLength       =   25
      TabIndex        =   0
      Tag             =   "00-Search for Surname"
      Top             =   4380
      Width           =   1695
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   5775
      Width           =   10440
      _Version        =   65536
      _ExtentX        =   18415
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
      Begin VB.CommandButton cmdVerify 
         Appearance      =   0  'Flat
         Caption         =   "Verify"
         Height          =   375
         Left            =   8160
         TabIndex        =   19
         Tag             =   "Multiple Employee Edit"
         Top             =   150
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdFullTable 
         Appearance      =   0  'Flat
         Caption         =   "Multiple Candidate Edit"
         Height          =   375
         Left            =   4920
         TabIndex        =   18
         Tag             =   "Multiple Candidate Edit"
         Top             =   150
         Width           =   2235
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
         Height          =   375
         Left            =   3360
         TabIndex        =   10
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Tag             =   "Print the Employee Listing"
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Tag             =   "Select the Employee listed above"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   960
         TabIndex        =   8
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   9120
         TabIndex        =   6
         Tag             =   "Print the Employee Listing"
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8280
         Top             =   360
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
         Left            =   9960
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
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmSFFind.frx":0000
      Height          =   4095
      Left            =   0
      OleObjectBlob   =   "frmSFFind.frx":0014
      TabIndex        =   11
      Top             =   120
      Width           =   10215
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
      Left            =   480
      TabIndex        =   4
      Top             =   4440
      Width           =   1665
   End
End
Attribute VB_Name = "frmSFFind"
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
Dim strNoAccessForms As String

Private Sub chkMassDele_Click()
    If chkMassDele.Value Then
        fraMassDele.Visible = True
    Else
        fraMassDele.Visible = False
    End If
End Sub

Private Sub cmdCancel_Click()

glbEEOK = False
'glbLEE_FName = "None Selected"
'glbLEE_SName = "None Selected"
'MsgBox Str(glbLEE_ID)
'Unload Me
glbUserUploadMode = UploadFormWithoutCheck: Unload Me
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim a%, Msg
    Dim SQLQ As String

    On Error GoTo DelErr
    
    If Data1.Recordset.EOF Then
        MsgBox "There is no record to be deleted."
        Exit Sub
    End If
    If Data1.Recordset("SF_UPT_DEMO") Or Data1.Recordset("SF_UPT_STATUS") Or Data1.Recordset("SF_UPT_POSITION") Or Data1.Recordset("SF_UPT_SALARY") Or Data1.Recordset("SF_UPT_REHIRE") Then
        MsgBox "Cannot delete this record because some update flags are Yes ."
        Exit Sub
    End If
    Msg = "Are You Sure You Want To Delete "
    Msg = Msg & "This Candidate Record?"
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub

    SQLQ = "DELETE FROM HRSF_XML_IMPORT  WHERE SF_CANDIDATE = " & Data1.Recordset("SF_CANDIDATE") & " "
    gdbAdoIhr001.Execute SQLQ
    
    'Data1.Recordset.Delete

    Data1.Refresh
    
    'Set FRS = Data1.Recordset.Clone
    'vbxTrueGrid.FetchRowStyle = True


Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRSF_XML_IMPORT", "Delete")
Call RollBack '10June99 js

End Sub

Private Sub cmdEESort_Click()
Dim xStr
txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Candidate List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

xFlagSortBySel = False
EEPIDSort = False

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by Candidate #"
    cmdEESort.Caption = "Sort by Surname "
    glbSort = "NUMBER"
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search by Surname"
    cmdEESort.Caption = "Sort by Candidate #"
    glbSort = "NAME"
End If

vbxTrueGrid.Tag = "ASC"

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
        'If EEPIDSort Then
        '    SQLQ = "ED_PAYROLL_ID  >= '" & Sch & "'"
        'Else
            If EESNameSort = True Then
                SQLQ = "SF_SURNAME  >= '" & Sch & "'"
                'Find this Employee# and then to search it. In order to find the "O'R" as Surname problem #7994
                xStr = "SELECT DISTINCT SF_CANDIDATE,SF_EMPNBR,SF_SURNAME,SF_FNAME,SF_PLANT,SF_HIRETYPE,SF_STARTDATE FROM HRSF_XML_IMPORT WHERE " & SQLQ
                xStr = xStr & " ORDER BY SF_SURNAME, SF_FNAME"
                rsEmpNum.Open xStr, gdbAdoIhr001, adOpenStatic
                If Not rsEmpNum.EOF Then
                    xSelect = "SF_CANDIDATE=" & rsEmpNum("SF_CANDIDATE")
                Else
                    xSelect = "SF_CANDIDATE=0" '" (1=1)"
                End If
                rsEmpNum.Close
                SQLQ = xSelect
                'Find this Employee# - End
            Else
                If Not IsNumeric(txtEESearch) Then
                    Beep
                    MsgBox "Employee Identification must be numeric"
                    Exit Sub
                End If
                SQLQ = "SF_CANDIDATE >= " & Sch & ""
            End If
        'End If
    Else
        If SortByFieldName = "SF_STARTDATE" Then 'For Date Field
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

Private Sub cmdFullTable_Click()
    vbxTrueGrid.AllowUpdate = True
    vbxTrueGrid.Enabled = True
    vbxTrueGrid.EditActive = True
    vbxTrueGrid.Refresh
    vbxTrueGrid.MarqueeStyle = 2
End Sub

Private Sub cmdOK_Click()
Dim xTmpFlag As Boolean
Dim X
' set the global last EE number for returned value
glbEEOK = True
If Data1.Recordset.EOF And Data1.Recordset.BOF Then 'laura 03/04/98
    Exit Sub
End If
    
glbCandidate = Data1.Recordset("SF_CANDIDATE")
glbCand_SF_ID = Data1.Recordset("SF_ID")

'Call ReDisplayFormsHRSoft

'Unload Me
If Not Data1.Recordset.EOF Then
    'Ticket #25911 Franks 03/17/2015 - begin
    xTmpFlag = True
    If IsNull(Data1.Recordset("SF_POSITIONCODE")) Then
        xTmpFlag = False
    Else
        If Len(Trim(Data1.Recordset("SF_POSITIONCODE"))) = 0 Then
            xTmpFlag = False
        End If
    End If
    If Not xTmpFlag Then
        MsgBox "No Position Code. Please fix it and then try it"
        Exit Sub
    End If
    'Ticket #25911 Franks 03/17/2015 - end
    
    Call HRSoftAction(Data1.Recordset)
    If glbCandidate = 0 Then
        Data1.Refresh
        If Not Data1.Recordset.EOF Then
            glbCandidate = Data1.Recordset("SF_CANDIDATE")
            glbCand_SF_ID = Data1.Recordset("SF_ID")
        End If
    End If
End If
End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String
Dim dscGroup$
'    cmdPrint.Enabled = False
'    glbiOneWhere = False
'    glbstrSelCri = ""
'    Call glbCri_DeptUN("")
'    If Len(glbstrSelCri) >= 0 Then
'        Me.vbxCrystal.SelectionFormula = glbstrSelCri
'    End If
'    RHeading = "Employee Listing"
'    Me.vbxCrystal.WindowTitle = RHeading & " Report"
'    Me.vbxCrystal.BoundReportHeading = RHeading
'    If Not xFlagSortBySel Then
'        If EEPIDSort Then
'            Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HREMP.ED_PAYROLL_ID};ANYCHANGE;" & Left(vbxTrueGrid.Tag, 1) 'A"
'        Else
'            If lblSearchBy.Caption = "Search by Surname" Then
'                Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;" & Left(vbxTrueGrid.Tag, 1) 'A"
'            Else
'                Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HREMP.ED_EMPNBR};ANYCHANGE;" & Left(vbxTrueGrid.Tag, 1) 'A"
'            End If
'        End If
'    Else
'        Me.vbxCrystal.GroupCondition(0) = "GROUP1;{HREMP." & UCase(SortByFieldName) & "};ANYCHANGE;" & Left(vbxTrueGrid.Tag, 1) 'A"
'    End If
'    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfind.rpt"
'    dscGroup$ = "PgHeading" & "= '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
'    Me.vbxCrystal.Formulas(0) = dscGroup$
'    If glbSQL Or glbOracle Then
'        Me.vbxCrystal.Connect = RptODBC_SQL
'    Else
'        Me.vbxCrystal.Connect = "PWD=petman;"
'        Me.vbxCrystal.DataFiles(0) = glbIHRDB
'    End If
'    Me.vbxCrystal.Action = 1
'    cmdPrint.Enabled = True
End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function EEList()
Dim SQLQ As String
Dim countr As Integer  ' EEList_Snap is definded at form level
Dim locSeleDeptUn

On Error GoTo EEList_Err
EEList = False
'SQLQ = "SELECT SF_ID, SF_CANDIDATE,SF_EMPNBR,SF_SURNAME,SF_FNAME,SF_PLANT,SF_HIRETYPE,SF_STARTDATE,SF_UPT_DEMO,SF_UPT_STATUS,SF_UPT_POSITION,SF_UPT_SALARY FROM HRSF_XML_IMPORT "
SQLQ = "SELECT * FROM HRSF_XML_IMPORT "
locSeleDeptUn = Replace(glbSeleDeptUn, "ED_", "SF_")
SQLQ = SQLQ & " WHERE " & locSeleDeptUn
'SQLQ = SQLQ & " AND ((SF_UPT_PROCESSED IS NULL) OR SF_UPT_PROCESSED = 0) "
SQLQ = SQLQ & " AND (SF_UPT_PROCESSED = 0) "
If Not xFlagSortBySel Then
    If cmdEESort.Caption = "Sort by Candidate #" Then '"Sort by Candidate Number" Then
        SQLQ = SQLQ & " ORDER BY SF_SURNAME, SF_FNAME"
    Else
        SQLQ = SQLQ & " ORDER BY SF_CANDIDATE"
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
Case 1
    xTitle = "Search by Surname" '"Search by Surname"
    EESNameSort = True
    'lblSearchBy.Caption = "Search by Surname"
    cmdEESort.Caption = "Sort by Employee Number"
    glbSort = "NAME"
    xFlagSortBySel = False
Case 3
    xTitle = "Search by Candidate #" '"Sort by Employee Number"
    EESNameSort = False
    'lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort.Caption = "Sort by Surname "
    glbSort = "NUMBER"
    xFlagSortBySel = False
Case Else
    xTitle = "Search by " & Me.vbxTrueGrid.Columns(xNo).Caption
    SortByFieldName = Me.vbxTrueGrid.Columns(xNo).DataField
    xFlagSortBySel = True
End Select
lblSearchBy.Caption = xTitle

vbxTrueGrid.Tag = "ASC"

If EEList() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = lblSearchBy.Caption '"Search by Surname "    'laura jan 05 1998

End Sub

Private Sub cmdStart_Click()
Dim SQLQ As String
Dim I As Integer
Dim Msg As String
Dim a%
    If Len(dlpDateRange(0).Text) = 0 Then
        MsgBox "From Date is a required field"
        dlpDateRange(0).SetFocus
        Exit Sub
    Else
        If Not IsDate(dlpDateRange(0).Text) Then
            MsgBox "Invalid From Date"
            dlpDateRange(0).SetFocus
            Exit Sub
        End If
    End If
    If Len(dlpDateRange(1).Text) = 0 Then
        MsgBox "To Date is a required field"
        dlpDateRange(1).SetFocus
        Exit Sub
    Else
        If Not IsDate(dlpDateRange(1).Text) Then
            MsgBox "Invalid To Date"
            dlpDateRange(1).SetFocus
            Exit Sub
        End If
    End If
    
    Msg = "All records with Processed checkboxes checked will be deleted. "
    Msg = Msg & Chr(10) & "Are you sure you want to do this?"
    Msg = Msg & "This Record?"
    a% = MsgBox(Msg, 36, "Confirm Delete")
    If a% <> 6 Then Exit Sub
    SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE SF_UPT_PROCESSED = 1 "
    SQLQ = SQLQ & "AND SF_FILEDATE >= " & Date_SQL(dlpDateRange(0).Text) & " "
    SQLQ = SQLQ & "AND SF_FILEDATE <= " & Date_SQL(dlpDateRange(1).Text) & " "
    gdbAdoIhr001.Execute SQLQ, I
    MsgBox I & " record(s) deleted."
End Sub

Private Sub cmdVerify_Click()
'Dim xHIRETYPE
'xHIRETYPE = ""
'If Not IsNull(Data1.Recordset("SF_HIRETYPE")) Then
'    xHIRETYPE = Data1.Recordset("SF_HIRETYPE")
'End If
'
'If xHIRETYPE = "REH" Then
'
'End If

'meeting with Jerry on 12/03/2013, don't work on Verify now because the "Multiple Candidate Edit" can change the data directly
End Sub

Private Sub Form_Activate()
glbOnTop = "frmSFFind"
Me.Height = 7005 '6705
Me.Width = 10545
'txtEESearch.SetFocus

End Sub

Private Sub Form_Load()
Dim xStr

glbOnTop = "frmSFFind"

'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB
'Data1.RecordSource = "HREMP"
If glbSort = "NUMBER" Then
    EESNameSort = True  'first sort is by surname
Else
    EESNameSort = False
End If

EEPIDSort = False

'
'If glbLinamar Then
'    Me.vbxTrueGrid.Columns(7).Visible = True
'    Me.vbxTrueGrid.Columns(7).Caption = lStr("Region")
'    Me.vbxTrueGrid.Columns(8).Caption = "Home Line"
'    Me.vbxTrueGrid.Columns(8).DataField = "ED_HOMELINE"
'Else
'    Me.vbxTrueGrid.Columns(7).Visible = False
'End If

'If gSec_Show_SIN_SSN Then
'    If Not glbLinamar Then
'        'Collectcorp Inc. asked this function
'        Me.vbxTrueGrid.Columns(8).Caption = "SIN/SSN"
'        Me.vbxTrueGrid.Columns(8).DataField = "ED_SIN"
'    End If
'End If

Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Retrieving Candidate List - Stand by"
Call cmdEESort_Click
'If glbLEE_SName = "Multi_EMP" Then
'    vbxTrueGrid.MultiSelect = 2
'    If glbLEE_FName <> "" Then
'        With Data1.Recordset
'            If Not .EOF Then .MoveLast
'            xStr = glbLEE_FName & ","
'            Do Until .BOF
'                If InStr(glbLEE_FName & ",", !EmpNbr & ",") <> 0 Then
'                    xStr = Replace(xStr, !EmpNbr & ",", "")
'                    vbxTrueGrid.SelBookmarks.Add vbxTrueGrid.Bookmark
'                    DoEvents
'                    If Trim(xStr) = "" Then Exit Do
'                End If
'                .MovePrevious
'            Loop
'        End With
'    End If
'End If

MDIMain.panHelp(0).Caption = "info:HR Main functions are Locked until EE is Selected"
Screen.MousePointer = DEFAULT
Me.vbxTrueGrid.Refresh

If Not gSec_Upd_Basic Then
    cmdDelete.Enabled = False
    chkMassDele.Enabled = False
    cmdFullTable.Enabled = False
End If

'If MDIMain.lstPanel.Visible = True Then
'
'End If

End Sub

Private Sub txtEESearch_GotFocus()
Call SetPanHelp(ActiveControl)

End Sub

Private Sub txtEESearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub vbxTrueGrid_DblClick()
'frmFind = True
'Call cmdOK_Click
'
'glbCandidate = Data1.Recordset("SF_CANDIDATE") ' new
'Call HRSoftAction(Data1.Recordset)
Call cmdOK_Click

End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
Dim locSeleDeptUn
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    'SQLQ = "SELECT DISTINCT SF_CANDIDATE,SF_EMPNBR,SF_SURNAME,SF_FNAME,SF_PLANT,SF_HIRETYPE,SF_STARTDATE FROM HRSF_XML_IMPORT "
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT "
    locSeleDeptUn = Replace(glbSeleDeptUn, "ED_", "SF_")
    SQLQ = SQLQ & " WHERE " & locSeleDeptUn
    SQLQ = SQLQ & " AND (SF_UPT_PROCESSED = 0) "
    SQLQ = SQLQ & " ORDER BY  " & UCase(vbxTrueGrid.Columns(ColIndex).DataField) & " " & vbxTrueGrid.Tag
    
    
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

Public Sub ReDisplayFormsHRSoft()
'    If glbOnTop = "frmHIRE" Then
'            'Call frmHIRE.Form_Load
'            Call frmHIRE.Display_Values
'    End If
End Sub
