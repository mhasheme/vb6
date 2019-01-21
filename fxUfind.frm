VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUFIND 
   Appearance      =   0  'Flat
   Caption         =   "Find Login User"
   ClientHeight    =   6000
   ClientLeft      =   660
   ClientTop       =   1050
   ClientWidth     =   10200
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
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   10200
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   5340
      Width           =   10200
      _Version        =   65536
      _ExtentX        =   17992
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   1890
         TabIndex        =   6
         Tag             =   "Print the Employee Listing"
         Top             =   150
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Tag             =   "Select the Employee listed above"
         Top             =   150
         Width           =   735
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
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxUfind.frx":0000
      Height          =   4605
      Left            =   120
      OleObjectBlob   =   "fxUfind.frx":0014
      TabIndex        =   3
      Tag             =   "Employee Listing "
      Top             =   90
      Width           =   9975
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Tag             =   "Find Employee"
      Top             =   4815
      Width           =   735
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "Sort by User ID"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   4815
      Width           =   2415
   End
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   25
      TabIndex        =   0
      Tag             =   "00-Search for Surname"
      Top             =   4860
      Width           =   1935
   End
   Begin VB.Label lblSearchBy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search by User Name"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   45
      TabIndex        =   7
      Top             =   4905
      Width           =   1860
   End
End
Attribute VB_Name = "frmUFIND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EEList_Snap As New ADODB.Recordset
Dim EESNameSort As Integer
'Dim OSN As Double, OSCh As String     ' last search items

Private Sub cmdCancel_Click()

glbEEOK = False
Unload Me

End Sub

Private Sub cmdEESort_Click()
Dim xStr

txtEESearch.Text = ""

Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing User List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by User ID"
    cmdEESort.Caption = "Sort by User Name "
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search User Name "
    cmdEESort.Caption = "Sort by User ID"
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

On Error GoTo Srch_Err

Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch, "'", "''")
    If EESNameSort = True Then
        SQLQ = "USERNAME  >= '" & Sch & "'"
    Else
        SQLQ = "USERID >= '" & Replace(Sch, "'", "''") & "'"
    End If
    Data1.Recordset.Find SQLQ
End If

If Data1.Recordset.EOF Then
    If Data1.Recordset.RecordCount > 0 Then Data1.Recordset.MoveFirst
    MsgBox "Employee not found"
End If

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

If glbLUserNAME = "Multi_EMP" Then
    If vbxTrueGrid.SelBookmarks.count = 0 Then
        glbLUserID = Data1.Recordset!UserID
    Else
        If vbxTrueGrid.SelBookmarks.count > 1000 Then
            MsgBox vbxTrueGrid.SelBookmarks.count & " users are selected" + Chr(10) + " Please make that less than 1000 users"
            Exit Sub
        End If
        glbLUserID = ""
        For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
            vbxTrueGrid.Bookmark = vbxTrueGrid.SelBookmarks(X)
            glbLUserID = glbLUserID & Data1.Recordset!UserID & ","
        Next
        glbLUserID = Left(glbLUserID, Len(glbLUserID) - 1)
    End If
Else
    glbLUserID = Data1.Recordset("USERID")
    glbLUserNAME = Data1.Recordset("USERNAME")
End If

Unload Me

End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdPrint_Click()
'Dim RHeading As String
'Dim dscGroup$
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
'    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgfind.rpt"
'    dscGroup$ = "PgHeading" & "= '" & Replace(RHeading, "'", "' + chr(39) + '") & "'"
'    Me.vbxCrystal.Formulas(0) = dscGroup$
'    If glbSQL Then
'        Me.vbxCrystal.Connect = RptODBC_SQL
'    Else
'        Me.vbxCrystal.Connect = "PWD=petman;"
'        Me.vbxCrystal.DataFiles(0) = glbIHRDB
'    End If
'    Me.vbxCrystal.Action = 1
'    cmdPrint.Enabled = True
'End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function EEList()
Dim SQLQ As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

On Error GoTo EEList_Err

EEList = False         ' if not found - no depts
SQLQ = "Select "
If glbLinamar Then
    SQLQ = SQLQ & "right(EMPNBR,3)+'-'+ left(EMPNBR,LEN(EMPNBR)-3) AS SHOW_EMPNBR,"
Else
    If glbOracle Then
        SQLQ = SQLQ & "EMPNBR AS SHOW_EMPNBR,"
    Else
        SQLQ = SQLQ & "STR(EMPNBR) AS SHOW_EMPNBR,"
    End If
    
End If
SQLQ = SQLQ & "EMPNBR,USERID, USERNAME,SECURE_TEMPLATE From HR_SECURE_BASIC"
SQLQ = SQLQ & " Where EMPNBR IS NULL OR EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")"

If cmdEESort.Caption = "Sort by User ID" Then
    SQLQ = SQLQ & " ORDER BY USERNAME"
Else
    SQLQ = SQLQ & " ORDER BY USERID"
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
Data1.ConnectionString = glbAdoIHRDB

If glbSort = "NUMBER" Then
    EESNameSort = True  'first sort is by surname
Else
    EESNameSort = False
End If


If glbLinamar Then vbxTrueGrid.Columns(2).DataField = "SHOW_EMPNBR"

Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Retrieving User List - Stand by"

Call cmdEESort_Click

If glbLUserNAME = "Multi_EMP" Then
    vbxTrueGrid.MultiSelect = 2
    If glbLUserID <> "" Then
        With Data1.Recordset
            If Not .EOF Then .MoveLast
            xStr = glbLUserID & ","
            Do Until .BOF
                If InStr(glbLUserID & ",", !UserID & ",") <> 0 Then
                    xStr = Replace(xStr, !UserID & ",", "")
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
End Sub

Private Sub txtEESearch_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtEESearch_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub vbxTrueGrid_DblClick()
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
    
    SQLQ = "Select "
    If glbLinamar Then
        SQLQ = SQLQ & "right(EMPNBR,3)+'-'+ left(EMPNBR,LEN(EMPNBR)-3) AS SHOW_EMPNBR,"
    Else
        If glbOracle Then
            SQLQ = SQLQ & "EMPNBR AS SHOW_EMPNBR,"
        Else
            SQLQ = SQLQ & "STR(EMPNBR) AS SHOW_EMPNBR,"
        End If
        
    End If
    SQLQ = SQLQ & "EMPNBR,USERID, USERNAME, SECURE_TEMPLATE From HR_SECURE_BASIC"
    SQLQ = SQLQ & " Where EMPNBR IS NULL OR EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & glbSeleDeptUn & ")"
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
    cmdOK.SetFocus
End If

End Sub

