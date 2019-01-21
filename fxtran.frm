VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmTranEMPL 
   Appearance      =   0  'Flat
   Caption         =   "Transfer Employee"
   ClientHeight    =   6000
   ClientLeft      =   660
   ClientTop       =   1050
   ClientWidth     =   11895
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
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6000
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   5
      Top             =   5340
      Width           =   11895
      _Version        =   65536
      _ExtentX        =   20981
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
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Tag             =   "Select the Employee listed above"
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   1905
         TabIndex        =   8
         Tag             =   "Print the Employee Listing"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   960
         TabIndex        =   7
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   825
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   120
         TabIndex        =   6
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
         ReportSource    =   1
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Tag             =   "Find Employee"
      Top             =   4860
      Width           =   735
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "Sort by Employee Number"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   4860
      Width           =   2415
   End
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "00-Search for Surname"
      Top             =   4860
      Width           =   1935
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fxtran.frx":0000
      Height          =   4635
      Left            =   0
      OleObjectBlob   =   "fxtran.frx":0014
      TabIndex        =   0
      Tag             =   "Employee Listing "
      Top             =   60
      Width           =   11880
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
      TabIndex        =   4
      Top             =   4860
      Width           =   1665
   End
End
Attribute VB_Name = "frmTranEMPL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EEList_Snap As New ADODB.Recordset
Dim EESNameSort As Integer
Private Sub cmdCancel_Click()
glbEEOK = False
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim Msg$, Title$, DgDef, Response%
Msg$ = Msg$ & Chr(10) & "Are you sure you wish to delete "
Msg$ = Msg$ & Chr(10) & "the transfer record?"
Title$ = "Delete Transfer Record"
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If
Screen.MousePointer = HOURGLASS
gdbAdoIhr001.Execute "UPDATE LN_TRALOG SET TL_TCOMPLETE='D' WHERE TL_ID=" & Data1.Recordset("TL_ID")
gdbAdoIhr001.Execute "DELETE FROM HR_PHOTO WHERE PT_EMPNBR=" & Data1.Recordset("ED_EMPNBR")
Data1.Refresh
Screen.MousePointer = DEFAULT
End Sub

Private Sub cmdEESort_Click()

txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

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
On Error GoTo Srch_Err
Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch, "'", "''")
    If EESNameSort = True Then
        'SQLQ = "ED_SURNAME  >= '" & Sch & "'"
        SQLQ = "ED_SURNAME  like '" & Sch & "%'"
    Else
        If Not glbLinamar And Not IsNumeric(txtEESearch) Then
            Beep
            MsgBox "Employee Identification must be numeric"
            Exit Sub
        End If
        If glbLinamar Then
            'SQLQ = "EMPNBR >= '" & Sch & "'"
            SQLQ = "EMPNBR like '" & Sch & "%'"
        Else
            SQLQ = "ED_EMPNBR >= '" & Sch & "'"
        End If
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
' set the global last EE number for returned value
glbEEOK = True
If Data1.Recordset.EOF And Data1.Recordset.BOF Then 'laura 03/04/98
    Exit Sub
End If
glbTran_ID = Data1.Recordset("ED_EMPNBR")
glbTran_Seq = Data1.Recordset("TL_TERM_SEQ")
glbTERM_Seq = glbTran_Seq
glbTERM_ID = glbTran_ID
If Not IsNull(Data1.Recordset("ED_FNAME")) Then
    glbTran_Fname = Data1.Recordset("ED_FNAME")
Else
    glbTran_Fname = "*ERROR*"
End If
If Not IsNull(Data1.Recordset("ED_SURNAME")) Then
    glbTran_SName = Data1.Recordset("ED_SURNAME")
Else
    glbTran_SName = "*ERROR*"
End If

Unload Me

End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Employee Listing"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Action = 1
End Sub

Private Sub cmdPrint_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function EEList()
Dim SQLQ As String
Dim countr   As Integer  ' EEList_Snap is definded at form level

On Error GoTo EEList_Err
EEList = False         ' if not found - no depts
If glbSamuel Then 'Ticket #20884 Franks 10/20/2011
    SQLQ = "Select  "
    SQLQ = SQLQ & "TL_NEWEMPNBR AS EMPNBR,"
    SQLQ = SQLQ & "TL_NEWEMPNBR AS ED_EMPNBR,"
    SQLQ = SQLQ & "TL_SURNAME AS ED_SURNAME,TL_FNAME AS ED_FNAME,"
    SQLQ = SQLQ & "TL_NEWDIVEDATE AS TDATE,"
    SQLQ = SQLQ & "TL_ID,TL_TERM_SEQ,"
    
    SQLQ = SQLQ & "(CASE WHEN V1.TB_DESC IS NULL THEN TL_OLDPLANT ELSE V1.TB_DESC END) AS FROMDIV,"
    SQLQ = SQLQ & "(CASE WHEN T.TB_DESC IS NULL THEN TL_TOREASON ELSE T.TB_DESC END) AS TREASON, "
    SQLQ = SQLQ & "(CASE WHEN V.TB_DESC IS NULL THEN TL_NEWPLANT ELSE V.TB_DESC END) AS TODIV "
    SQLQ = SQLQ & "FROM LN_TRALOG A "
    SQLQ = SQLQ & "LEFT JOIN HRTABL T ON A.TL_TOREASON_TABL=T.TB_NAME AND A.TL_TOREASON=T.TB_KEY "
    SQLQ = SQLQ & "LEFT JOIN HRTABL V ON A.TL_NEWPLANT_TABL=V.TB_NAME AND A.TL_NEWPLANT=V.TB_KEY "
    SQLQ = SQLQ & "LEFT JOIN HRTABL V1 ON A.TL_OLDPLANT_TABL=V1.TB_NAME AND A.TL_OLDPLANT=V1.TB_KEY "
    SQLQ = SQLQ & "WHERE TL_TCOMPLETE='N' AND TL_TYPE='TOUT' "
    SQLQ = SQLQ & "AND TL_NEWPLANT IN (SELECT TB_KEY FROM HRTABL WHERE TB_NAME = 'EDAB' AND " & glbSeleAdminBy & ") "
    If cmdEESort.Caption = "Sort by Employee Number" Then
        SQLQ = SQLQ & " ORDER BY TODIV,ED_SURNAME, ED_FNAME"
    Else
        If glbLinamar Then
            SQLQ = SQLQ & " ORDER BY TODIV,EMPNBR"
        Else
            SQLQ = SQLQ & " ORDER BY TODIV,ED_EMPNBR"
        End If
    End If
Else 'Linamar and WFC
    SQLQ = "Select  "
    If glbLinamar Then
    SQLQ = SQLQ & "right(TL_NEWEMPNBR,3)+'-'+ left(TL_NEWEMPNBR,LEN(TL_NEWEMPNBR)-3) AS EMPNBR,"
    End If
    If glbWFC Then
        SQLQ = SQLQ & "TL_NEWEMPNBR AS EMPNBR,"
    End If
    SQLQ = SQLQ & "TL_NEWEMPNBR AS ED_EMPNBR,"
    
    SQLQ = SQLQ & "TL_SURNAME AS ED_SURNAME,TL_FNAME AS ED_FNAME,"
    
    SQLQ = SQLQ & "TL_NEWDIVEDATE AS TDATE,"
    SQLQ = SQLQ & "TL_ID,TL_TERM_SEQ,"
    
    'Ticket #21677 Franks 03/14/2012
    If glbWFC Then
        SQLQ = SQLQ & "TL_OLD_ORG,TL_NEW_ORG,TL_OLD_ORG_DESC,TL_NEW_ORG_DESC,"
    End If
    
    SQLQ = SQLQ & "(CASE WHEN V1.DIVISION_NAME IS NULL THEN TL_NEWDIV ELSE V1.DIVISION_NAME END) AS FROMDIV,"
    SQLQ = SQLQ & "(CASE WHEN TB_DESC IS NULL THEN TL_TOREASON ELSE TB_DESC END) AS TREASON, "
    SQLQ = SQLQ & "(CASE WHEN V.DIVISION_NAME IS NULL THEN TL_NEWDIV ELSE V.DIVISION_NAME END) AS TODIV "
    SQLQ = SQLQ & "FROM LN_TRALOG A "
    SQLQ = SQLQ & "LEFT JOIN HRTABL T ON A.TL_TOREASON_TABL=T.TB_NAME AND A.TL_TOREASON=T.TB_KEY "
    SQLQ = SQLQ & "LEFT JOIN HR_DIVISION V ON A.TL_NEWDIV=V.DIV "
    SQLQ = SQLQ & "LEFT JOIN HR_DIVISION V1 ON A.TL_OLDDIV=V1.DIV "
    SQLQ = SQLQ & "WHERE TL_TCOMPLETE='N' AND TL_TYPE='TOUT' "
    SQLQ = SQLQ & "AND TL_NEWDIV IN (SELECT DIV FROM HR_DIVISION WHERE " & glbSeleDiv & ") "
    If cmdEESort.Caption = "Sort by Employee Number" Then
        SQLQ = SQLQ & " ORDER BY TODIV,ED_SURNAME, ED_FNAME"
    Else
        If glbLinamar Then
            SQLQ = SQLQ & " ORDER BY TODIV,EMPNBR"
        Else
            SQLQ = SQLQ & " ORDER BY TODIV,ED_EMPNBR"
        End If
    End If
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

'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = glbAdoIHRDB
'Data1.RecordSource = "HREMP"
If glbSort = "NUMBER" Then
    EESNameSort = True  'first sort is by surname
Else
    EESNameSort = False
End If

'frmTranEMPL.Top = 1365
'frmTranEMPL.Left = 1140


Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Retrieving Employee List - Stand by"
'If EEList() = False Then     ' get the info for this person
'    Exit Sub
'End If          ' dpartment specific and populate the list
If glbWFC Then
    vbxTrueGrid.Columns(0).Caption = lStr("From Division")
    vbxTrueGrid.Columns(1).Caption = lStr("To Division")
Else
    vbxTrueGrid.Columns(2).Visible = False
    vbxTrueGrid.Columns(3).Visible = False
End If
If glbSamuel Then
    vbxTrueGrid.Columns(0).Caption = "From " & lStr("Administered By")
    vbxTrueGrid.Columns(1).Caption = "To " & lStr("Administered By")
End If
MDIMain.panHelp(0).Caption = "info:HR Main functions Locked until EE Selected"
Screen.MousePointer = DEFAULT
Call cmdEESort_Click
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
        
        SQLQ = "Select  "
        If glbLinamar Then
        SQLQ = SQLQ & "right(TL_NEWEMPNBR,3)+'-'+ left(TL_NEWEMPNBR,LEN(TL_NEWEMPNBR)-3) AS EMPNBR,"
        End If
        If glbWFC Then
            SQLQ = SQLQ & "TL_NEWEMPNBR AS EMPNBR,"
        End If
        SQLQ = SQLQ & "TL_NEWEMPNBR AS ED_EMPNBR,"
        
        SQLQ = SQLQ & "TL_SURNAME AS ED_SURNAME,TL_FNAME AS ED_FNAME,"
        
        SQLQ = SQLQ & "TL_NEWDIVEDATE AS TDATE,"
        SQLQ = SQLQ & "TL_ID,TL_TERM_SEQ,"
        
        SQLQ = SQLQ & "(CASE WHEN V1.DIVISION_NAME IS NULL THEN TL_NEWDIV ELSE V1.DIVISION_NAME END) AS FROMDIV,"
        SQLQ = SQLQ & "(CASE WHEN TB_DESC IS NULL THEN TL_TOREASON ELSE TB_DESC END) AS TREASON, "
        SQLQ = SQLQ & "(CASE WHEN V.DIVISION_NAME IS NULL THEN TL_NEWDIV ELSE V.DIVISION_NAME END) AS TODIV "
        SQLQ = SQLQ & "FROM LN_TRALOG A "
        SQLQ = SQLQ & "LEFT JOIN HRTABL T ON A.TL_TOREASON_TABL=T.TB_NAME AND A.TL_TOREASON=T.TB_KEY "
        SQLQ = SQLQ & "LEFT JOIN HR_DIVISION V ON A.TL_NEWDIV=V.DIV "
        SQLQ = SQLQ & "LEFT JOIN HR_DIVISION V1 ON A.TL_OLDDIV=V1.DIV "
        SQLQ = SQLQ & "WHERE TL_TCOMPLETE='N' AND TL_TYPE='TOUT' "
        SQLQ = SQLQ & "AND TL_NEWDIV IN (SELECT DIV FROM HR_DIVISION WHERE " & glbSeleDiv & ") "
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





