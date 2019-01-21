VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmTLAY 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Temporary Lay Off Emplyees"
   ClientHeight    =   8595
   ClientLeft      =   480
   ClientTop       =   1140
   ClientWidth     =   11880
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame frmWFCBenList 
      Height          =   1335
      Left            =   9600
      TabIndex        =   37
      Top             =   5040
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CheckBox chkAllDates 
         Caption         =   "All Date"
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
         Left            =   4560
         TabIndex        =   38
         Top             =   2445
         Width           =   1395
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid1 
         Bindings        =   "frmTLAY.frx":0000
         Height          =   2025
         Left            =   120
         OleObjectBlob   =   "frmTLAY.frx":0014
         TabIndex        =   39
         Top             =   240
         Width           =   10275
      End
      Begin INFOHR_Controls.DateLookup dlpEndDate 
         DataSource      =   "Data2"
         Height          =   285
         Left            =   1800
         TabIndex        =   40
         Tag             =   "41-Effective date of salary change"
         Top             =   2400
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpLastDate 
         Height          =   285
         Left            =   1800
         TabIndex        =   41
         Tag             =   "41-Effective date of salary change"
         Top             =   2760
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Day"
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
         Left            =   360
         TabIndex        =   43
         Top             =   2820
         Width           =   1365
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
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
         Index           =   7
         Left            =   360
         TabIndex        =   42
         Top             =   2460
         Width           =   885
      End
   End
   Begin VB.Frame fraReActivate 
      Caption         =   "Re-activate"
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   5880
      Visible         =   0   'False
      Width           =   9075
      Begin INFOHR_Controls.DateLookup dlpUnion 
         DataSource      =   " "
         Height          =   285
         Left            =   6480
         TabIndex        =   15
         Tag             =   "40-Union Date"
         Top             =   810
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.ComboBox comESalInc 
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
         Left            =   2660
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Tag             =   "Eligible for Salary Increase"
         Top             =   1260
         Visible         =   0   'False
         Width           =   975
      End
      Begin INFOHR_Controls.DateLookup dlpEDate 
         Height          =   285
         Index           =   1
         Left            =   6360
         TabIndex        =   14
         Tag             =   "41-Effective to Date"
         Top             =   810
         Visible         =   0   'False
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpEDate 
         Height          =   285
         Index           =   0
         Left            =   2340
         TabIndex        =   11
         Tag             =   "41-Effecive Date"
         Top             =   810
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   10
         Tag             =   "01-Employment Status"
         Top             =   360
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin VB.CheckBox chkLeave 
         Caption         =   "Leave"
         Height          =   195
         Left            =   6240
         TabIndex        =   31
         Top             =   120
         Visible         =   0   'False
         Width           =   1515
      End
      Begin INFOHR_Controls.DateLookup dlpDOther1 
         DataSource      =   " "
         Height          =   285
         Left            =   7320
         TabIndex        =   13
         Tag             =   "40-Other Date 2"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Eligible for Salary Increase"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   36
         Top             =   1260
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Label lbOtherDate1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Date 1"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5640
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   5880
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Effective Date As Of"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   870
         Width           =   1770
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "New Employment Status"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   390
         Width           =   2055
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmTLAY.frx":4C49
      Height          =   2415
      Left            =   0
      OleObjectBlob   =   "frmTLAY.frx":4C5D
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "&Sort by Emp #"
      Height          =   375
      Index           =   0
      Left            =   6780
      TabIndex        =   3
      Tag             =   "Change the sorting method of the Employee List"
      Top             =   2820
      Width           =   2475
   End
   Begin VB.CommandButton cmdEESort 
      Appearance      =   0  'Flat
      Caption         =   "&Sort by Surname"
      Height          =   375
      Index           =   1
      Left            =   6780
      TabIndex        =   24
      Top             =   2820
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.TextBox txtEESearch 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2220
      TabIndex        =   1
      Tag             =   "00-Search for Surname"
      Top             =   2850
      Width           =   1935
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      Caption         =   "&Find"
      Height          =   375
      Left            =   4260
      TabIndex        =   2
      Tag             =   "Find Employee"
      Top             =   2820
      Width           =   735
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   180
      Left            =   0
      TabIndex        =   17
      Top             =   8415
      Visible         =   0   'False
      Width           =   11880
      _Version        =   65536
      _ExtentX        =   20955
      _ExtentY        =   317
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
         Height          =   405
         Left            =   6960
         Top             =   180
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
      End
   End
   Begin VB.Frame fraTerminate 
      Caption         =   "Terminamte (Will Go To Terminamtion Screen)"
      Height          =   1275
      Left            =   120
      TabIndex        =   23
      Top             =   7440
      Visible         =   0   'False
      Width           =   9075
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Left            =   2520
         TabIndex        =   32
         Top             =   720
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   16
         Tag             =   "01-Termination Code - Code "
         Top             =   360
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   29
         Top             =   360
         Width           =   1710
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   28
         Tag             =   "41-Date Terminated"
         Top             =   750
         Width           =   1470
      End
   End
   Begin VB.Frame fraOptions 
      Height          =   1575
      Left            =   120
      TabIndex        =   19
      Top             =   3420
      Width           =   6255
      Begin VB.OptionButton Options 
         Caption         =   "Transfer Out"
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   7
         Tag             =   "Terminate"
         Top             =   1140
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton Options 
         Caption         =   "Terminate"
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Tag             =   "Terminate"
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Options 
         Caption         =   "Re-activate"
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Tag             =   "40-Re-activate"
         Top             =   540
         Value           =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Options 
         Caption         =   "Extending"
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Tag             =   "Extending"
         Top             =   240
         Width           =   2715
      End
   End
   Begin VB.Frame fraExtending 
      Caption         =   "Change Return Date"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   5070
      Visible         =   0   'False
      Width           =   9075
      Begin VB.CheckBox chkATPaidHours 
         Caption         =   "Paid Hours in AT"
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
         Left            =   6720
         TabIndex        =   33
         Top             =   270
         Visible         =   0   'False
         Width           =   1755
      End
      Begin INFOHR_Controls.DateLookup dlpTLAYDate 
         Height          =   285
         Left            =   3120
         TabIndex        =   8
         Tag             =   "41-Date Extending"
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDOther2 
         DataSource      =   " "
         Height          =   285
         Left            =   3120
         TabIndex        =   9
         Tag             =   "40-Other Date 2"
         Top             =   600
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   503
         ShowDescription =   0   'False
         TextBoxWidth    =   1180
      End
      Begin VB.Label lbOtherDate2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Other Date 2"
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
         Left            =   180
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Label lblWeeks 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0 Week"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   5160
         TabIndex        =   30
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "LOA Change  to "
         Height          =   375
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   2775
      End
   End
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   10440
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
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
   Begin VB.Label lblSearchBy 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Surname"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   25
      Top             =   2880
      Width           =   1665
   End
End
Attribute VB_Name = "frmTLAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EESNameSort
Dim fglbFollowID
Dim fdFdate, fdTdate, fglbNew
Dim fAttCode
Dim MailBody As String
Dim xNGStmpDate 'Ticket #19266
Dim xoLDAY 'Ticket #23920 Franks 07/04/2013

Private Sub clpCode_Change(Index As Integer)
If Index = 1 Then EMPCode_Desc
End Sub

Sub cmdClose_Click()
Unload Me

End Sub

'Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub cmdEESort_Click(Index As Integer)

txtEESearch.Text = ""
Screen.MousePointer = HOURGLASS
MDIMain.panHelp(0).Caption = "Refreshing Employee List - Stand by"
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "

If EESNameSort = True Then  ' was sorted by surname
    EESNameSort = False
    lblSearchBy.Caption = "Search by Emp. #"
    cmdEESort(0).Visible = False
    cmdEESort(1).Visible = True
Else
    EESNameSort = True
    lblSearchBy.Caption = "Search by Surname"
    cmdEESort(0).Visible = True
    cmdEESort(1).Visible = False
End If

If EEList() = 0 Then     ' get the info for this person
    Exit Sub
End If          ' dpartment specific and populate the list

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).Caption = " "
txtEESearch.SetFocus

End Sub

Private Sub cmdEESort_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdFind_Click()
Dim Sch As String, SQLQ As String
Dim bkmark

On Error GoTo Srch_Err

If Not Len(txtEESearch) > 0 Then
   MsgBox "To search you must enter something to search for."
   Exit Sub
End If
Data1.Refresh
If Not Data1.Recordset.EOF Then
    Sch = Replace(txtEESearch.Text, "'", "''")
    If EESNameSort = True Then
        SQLQ = "ED_SURNAME  >= '" & Sch & "'"
    Else
        If Not IsNumeric(txtEESearch.Text) And Not glbLinamar Then
            Beep
            MsgBox "Employee Identification must be numeric"
            Exit Sub
        End If
        If glbLinamar Then
            SQLQ = "EMPNBR >= '" & Sch & "'"
        Else
            SQLQ = "ED_EMPNBR >= '" & Sch & "'"
        End If

    End If
    Data1.Recordset.Find SQLQ
End If
If Data1.Recordset.EOF Then
    MsgBox "Employee not found"
    Data1.Refresh
End If

Exit Sub

Srch_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EEList", "HREMP", "Find Next")
Call RollBack '28July99 jsEnd Sub
End Sub

Private Sub cmdFind_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub cmdOK_Click()
Dim Msg$, DgDef As Variant, Response%
Dim Title$, EID&, TermDate$
Dim SQLQ
Dim xEMP
Dim SavEmp As String
Dim SavDOther 'Ticket #19266
Dim xPenStatus

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then Exit Sub

If Not chkTLAY() Then Exit Sub

glbLEE_ID = Data1.Recordset("ED_EMPNBR")

'This will resolve the error 3704 when the user first logs into info:HR and goes to LOA - Reactive and then
'reactivates and employee and then goes to Status/Dates screen then to Employee History. Since the global
'variable employee _FName and _SName were blank it was giving this error.
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

If Options(2) Or Options(3) Then
    If gSec_Inq_Terminations Then
        Screen.MousePointer = HOURGLASS
        Unload frmETERM
        glbTermTran = Options(2)
        
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
        frmETERM.dlpTermDate.Text = dlpTermDate.Text
        
        If glbLinamar Then
            glbLEE_ProdLine = Mid(Data1.Recordset("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", Data1.Recordset("PROD_LINE")) 'Ticket #14775
        End If
        
        If glbLinamar Then frmETERM.clpCode(1).Text = IIf(Options(2), "INLA", "TOUT")
        Load frmETERM
        Screen.MousePointer = DEFAULT
    Else
        MsgBox "You Do Not Have Authority For This Transaction"
    End If
    frmETERM.Show
    Exit Sub
End If

Msg$ = Msg$ & Chr(10) & "Are you sure you want to "
If Options(0) Then
    If glbLinamar Then
        Msg$ = Msg$ & "extend to 35 weeks " & Chr(10) & "this employee's Lay-Off?"
    Else
        Msg$ = Msg$ & "Change " & Chr(10) & "this employee's leave?"
    End If
End If
If Options(1) Then Msg$ = Msg$ & Options(1).Caption & Chr(10) & "this employee?"
'If Options(2) Then Msg$ = Msg$ & Options(2).Caption & Chr(10) & "this employee?"
'If Options(3) Then Msg$ = Msg$ & Options(3).Caption & Chr(10) & "this employee?"


If glbLinamar Then Title$ = "Temporary Lay Off" Else Title$ = "LOF Date Change"

DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
If Options(0) Or Options(1) Then Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.

If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If

MDIMain.panHelp(0).FloodType = 1

Screen.MousePointer = vbHourglass

If Options(0) Or Options(1) Then
    If Not AUDITTERM() Then MsgBox "ERROR - AUDIT FILE"
End If
Call updFollow
Call updStatus
If Not Options(1) Then
    Call updAttendance
End If

'Hemu - 18/08/2003 Begin - Jerry asked to remove the attendance records when the employee
'                          is reactivated from the Leave of Absence before the return date.
If Options(1) Then       'When Re-Activate

    SQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_DOA >= " & Date_SQL(dlpEDate(0))
    SQLQ = SQLQ & " AND AD_EMPNBR = " & Data1.Recordset!ED_EMPNBR
    SQLQ = SQLQ & " AND AD_REASON in (SELECT SC_ATTREASON FROM HRSTATUS "
    SQLQ = SQLQ & " WHERE SC_EMPNBR = " & Data1.Recordset!ED_EMPNBR & " AND SC_REASON = 'LOA' "
    SQLQ = SQLQ & " AND SC_FDATE = " & Date_SQL(Data1.Recordset(fdFdate)) & " AND SC_TDATE = "
    SQLQ = SQLQ & Date_SQL(Data1.Recordset(fdTdate)) & ")"
    
    gdbAdoIhr001.Execute SQLQ
       
End If
'Hemu - 18/08/2003 End

SQLQ = "UPDATE HREMP SET "
If Options(1) Then
    If glbWFC Then 'Ticket #25352 Franks 04/16/2014
    'Reactive from a Leave: "   If Status = SALC, remove the Last Day
        If Not IsNull(Data1.Recordset!ED_EMP) Then
            If Data1.Recordset!ED_EMP = "SALC" Then
                Data1.Recordset!ED_LDAY = Null
            End If
        End If
    End If
    
    If glbLinamar Then
        'Ticket #16456, Human Resource Report needs it
        'Data1.Recordset(fdFdate) = Null
        If IsDate(dlpEDate(0).Text) Then Data1.Recordset(fdTdate) = dlpEDate(0).Text
    Else
        If IsDate(dlpEDate(0).Text) Then Data1.Recordset(fdFdate) = dlpEDate(0).Text
        If glbCompSerial = "S/N - 2380W" Or glbCompSerial = "S/N - 2384W" Then
        'Vitalaire Ticket #12616
        'Ticket #20105 St. Marys Franks 10/11/2011
            If IsDate(dlpEDate(0).Text) Then
                Data1.Recordset("ED_LTHIRE") = dlpEDate(0).Text 'Audit
            End If
        End If
        If IsDate(dlpEDate(1).Text) Then Data1.Recordset(fdTdate) = dlpEDate(1).Text Else Data1.Recordset(fdTdate) = Null
        
        'Ticket #29985 - County of Essex - Enter the Union Date
        If glbCompSerial = "S/N - 2192W" Then
            If IsDate(dlpUnion.Text) Then Data1.Recordset("ED_UNION") = dlpUnion.Text Else Data1.Recordset("ED_UNION") = Null
        End If
    End If
    
    If Data1.Recordset!ED_EMP <> clpCode(1).Text Then
        If Len(clpCode(1).Text) > 0 Then xEMP = clpCode(1).Text Else xEMP = "*"
        If dlpEDate(1).Visible = True Then
            If IsDate(dlpEDate(1).Text) Then
                If Not EmpHisCalc(1, Data1.Recordset!ED_EMPNBR, "", "", xEMP, "", "", "", "", Date, , , dlpEDate(0).Text, dlpEDate(1).Text) Then MsgBox "EMPHIS Error"
            Else
                If Not EmpHisCalc(1, Data1.Recordset!ED_EMPNBR, "", "", xEMP, "", "", "", "", Date, , , dlpEDate(0).Text) Then MsgBox "EMPHIS Error"
            End If
        Else
            If Not EmpHisCalc(1, Data1.Recordset!ED_EMPNBR, "", "", xEMP, "", "", "", "", Date, , , dlpEDate(0).Text) Then MsgBox "EMPHIS Error"
        End If
    End If
    SavEmp = Data1.Recordset!ED_EMP
    Data1.Recordset!ED_EMP = clpCode(1).Text
End If

If Options(0) Then
    If glbCompSerial = "S/N - 2351W" Then 'Ticket #19858 - Burlington Technologies only
        'Ticket #14623 - Begin
        SQLQ = "UPDATE HR_JOB_HISTORY SET JH_ENDDATE = " & Date_SQL(dlpTLAYDate.Text) & " "
        SQLQ = SQLQ & "WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & Data1.Recordset!ED_EMPNBR
        gdbAdoIhr001.Execute SQLQ
        'Ticket #14623 - End
    End If
    SavDOther = dlpDOther2.Text 'Ticket #19266
    Data1.Recordset(fdTdate) = dlpTLAYDate.Text
    dlpTLAYDate.Text = ""

End If
Data1.Recordset.Update

'Ticket #30479 - Daily Entitlement - Recompute the Daily Accrual
If Options(1).Value Then
    If glbCompEntVacDaily Then
        Call Recompute_DailyAccrualFile(glbLEE_ID, dlpEDate(0).Text)
    End If
End If

'Samuel Ticket #20648 Franks 09/23/2011
If glbCompSerial = "S/N - 2382W" Then
    If Options(1).Value Then
        Call EmployeeFlagUpd(Data1.Recordset!ED_EMPNBR, 3, comESalInc.Text, dlpEDate(0).Text, "", "")
    End If
End If

'For WFC Pension System
If glbWFC Then
    If Options(1) Then
        toSOURCE = "IHR Re-Active from Leave" 'Ticket #19954
        'Ticket #21597 Franks 05/01/2012
        '"   Lookup the Table Master for the New Employment Status. If there is a Pension Status (tb_usr1) entered and is different than the current Pension Status, update the Pension Master with the Pension Status and Effective Date.
        xPenStatus = getPenStatusFromHRTABL(clpCode(1).Text)
        
        ''If SavEmp = "TLAY" Then 'Ticket #21788 Franks 03/26/2012
        'Ticket #21597 Franks 05/01/2012
        If SavEmp = "TLAY" Or Len(xPenStatus) > 0 Then 'Ticket #21788 Franks 03/26/2012
            If WFCPensionEligible(Data1.Recordset!ED_EMPNBR) Then
                Call uptEmpDates(Data1.Recordset!ED_EMPNBR, "ED_LDAY", Null)
                
                Call WFCPensionMasUpt(Data1.Recordset!ED_EMPNBR, "Re-Active from a Leave", dlpEDate(0).Text, xPenStatus, Year(CVDate(dlpEDate(0).Text)))
                'STD -> Others
                'If the Employment Status goes from STD to any status but LTD or DIS,
                'delete the Disability Date (ER_PENSIONDATE2)
                'If SavEmp = "STD" Then
                '    If SavEmp <> clpCode(1).Text Then
                '        If Not (clpCode(1).Text = "LTD" Or clpCode(1).Text = "DIS") Then
                '            Call Upt_PENSIONDATE2(Data1.Recordset!ED_EMPNBR, "DELETE")
                '        End If
                '    End If
                'End If
            End If
        End If
        
        'Ticket #19954 Franks 03/28/2011
        'Return from a Leave should delete the date in Pension Date 2
        Call Upt_PENSIONDATE2(Data1.Recordset!ED_EMPNBR, "DELETE")
        
        'Ticket #19266 Franks 11/30/2010
        If dlpDOther1.Visible Then
            SavDOther = dlpDOther1.Text
            Call WFC_NGS_Trans(Data1.Recordset!ED_EMPNBR, SavDOther, "Re-Active")
        End If
        
        'Ticket #23920 Franks 07/04/2013
        If frmWFCBenList.Visible Then 'US NGS employees
            Call WFC_NGSBenEndDateUpt(Data1.Recordset!ED_EMPNBR)
        End If
    End If
    If Options(0) Then
        'Ticket #19266 Franks 11/29/2010
        If dlpDOther2.Visible Then
            Call WFC_NGS_Trans(Data1.Recordset!ED_EMPNBR, SavDOther, "LOA Change")
        End If
    End If
End If

'For Linamar - Add a new Position Record with RTL as Reason for Change code
If Options(1) Then
    If glbSamuel Then 'Ticket #20885 Franks 12/01/2011
        SavDOther = dlpEDate(0).Text
        Call SAMUEL_Trans(Data1.Recordset!ED_EMPNBR, SavDOther, "Reture Leave")
    End If

    If glbLinamar Then
        Dim rsJOB As New ADODB.Recordset
        Dim rsSal As New ADODB.Recordset
        Dim rsJobDoc As New ADODB.Recordset
        Dim rJob, rSDate, rDHRS, rWHRS, rPHRS, rRepAut, rShift, rFTENum, rFTEHrs, rOrg, rPTFT, rComment, rComment2
        Dim rRepAut2, rRepAut3, rLeadHand, rLabourCD, rLabourDate, rUsrLabel, rUsrCheck, rUsrDate
        Dim rUsrLabel2, rUsrCheck2, rUsrDate2, rUsrLabel3, rUsrCheck3, rUsrDate3, rDiv, rDeptNo, rEmp
        Dim rGLNo, rSect, rRegion, rPosCtrl, rPayrollID, rGrid, rPayCateg, rBillRate, rEndDate, rEndReason
        Dim rSal, rSalCD, rGrade, rPayP, rReas1, rSalPC1, rSalChg1, rReas2, rSalPC2, rSalChg2, rReas3, rSalPC3, rSalChg3
        Dim rCompa, rNextDate, rSComment, rSComment2, rJobID
        
        'New Position Record
        rsJOB.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & Data1.Recordset!ED_EMPNBR, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        If Not rsJOB.EOF Then
            'Retrieve Data
            rJob = rsJOB("JH_JOB")
            rSDate = rsJOB("JH_SDATE")
            rDHRS = rsJOB("JH_DHRS")
            rWHRS = rsJOB("JH_WHRS")
            rPHRS = rsJOB("JH_PHRS")
            rRepAut = rsJOB("JH_REPTAU")
            rShift = rsJOB("JH_SHIFT")
            rFTENum = rsJOB("JH_FTENUM")
            rFTEHrs = rsJOB("JH_FTEHRS")
            rOrg = rsJOB("JH_ORG")
            rPTFT = rsJOB("JH_PT")
            rComment = rsJOB("JH_COMMENT")
            rComment2 = rsJOB("JH_COMMENT2")
            rRepAut2 = rsJOB("JH_REPTAU2")
            rRepAut3 = rsJOB("JH_REPTAU3")
            rLeadHand = rsJOB("JH_LEADHAND")
            rLabourCD = rsJOB("JH_LABOURCD")
            rLabourDate = rsJOB("JH_LABOUREDATE")
            rUsrLabel = rsJOB("JH_USRLABEL")
            rUsrCheck = rsJOB("JH_USRCHECK")
            rUsrDate = rsJOB("JH_USREDATE")
            rUsrLabel2 = rsJOB("JH_USRLABEL2")
            rUsrCheck2 = rsJOB("JH_USRCHECK2")
            rUsrDate2 = rsJOB("JH_USREDATE2")
            rUsrLabel3 = rsJOB("JH_USRLABEL3")
            rUsrCheck3 = rsJOB("JH_USRCHECK3")
            rUsrDate3 = rsJOB("JH_USREDATE3")
            rDiv = rsJOB("JH_DIV")
            rDeptNo = rsJOB("JH_DEPTNO")
            rEmp = rsJOB("JH_EMP")
            rGLNo = rsJOB("JH_GLNO")
            rSect = rsJOB("JH_SECTION")
            rRegion = rsJOB("JH_REGION")
            rPosCtrl = rsJOB("JH_POSITION_CONTROL")
            rPayrollID = rsJOB("JH_PAYROLL_ID")
            rGrid = rsJOB("JH_GRID")
            rPayCateg = rsJOB("JH_PAYROLL_CATEGORY")
            rBillRate = rsJOB("JH_BILLINGRATE")
            rEndDate = rsJOB("JH_ENDDATE")
            rEndReason = rsJOB("JH_ENDREAS")
                        
            'Remove the Current check from the existing current position record
            rsJOB("JH_CURRENT") = False
            rsJOB("JH_ENDDATE") = DateAdd("d", -1, dlpEDate(0).Text)
            rsJOB("JH_ENDREAS") = "TEMP"
            rsJOB("JH_LDATE") = Format(Now, "Short Date")
            rsJOB("JH_LTIME") = Time$
            rsJOB("JH_LUSER") = "999999999"
            rsJOB.Update
        
            'Add a new Position records with RTL reason for change
            rsJOB.AddNew
            rsJOB("JH_COMPNO") = "001"
            rsJOB("JH_EMPNBR") = Data1.Recordset!ED_EMPNBR
            rsJOB("JH_JOB") = rJob
            rsJOB("JH_SDATE") = dlpEDate(0).Text
            rsJOB("JH_DHRS") = rDHRS
            rsJOB("JH_WHRS") = rWHRS
            rsJOB("JH_PHRS") = rPHRS
            CheckHRTABLCode "SDRC", "RTL", "Returns from Temporary Layoff"
            rsJOB("JH_JREASON") = "RTL" 'Returns from Temporary Layoff
            rsJOB("JH_CURRENT") = True
            rsJOB("JH_REPTAU") = rRepAut
            rsJOB("JH_SHIFT") = rShift
            rsJOB("JH_FTENUM") = rFTENum
            rsJOB("JH_FTEHRS") = rFTEHrs
            rsJOB("JH_ORG") = rOrg
            rsJOB("JH_PT") = rPTFT
            rsJOB("JH_COMMENT") = rComment
            rsJOB("JH_COMMENT2") = rComment2
            rsJOB("JH_REPTAU2") = rRepAut2
            rsJOB("JH_REPTAU3") = rRepAut3
            rsJOB("JH_LEADHAND") = rLeadHand
            rsJOB("JH_LABOURCD") = rLabourCD
            rsJOB("JH_LABOUREDATE") = rLabourDate
            rsJOB("JH_USRLABEL") = rUsrLabel
            rsJOB("JH_USRCHECK") = rUsrCheck
            rsJOB("JH_USREDATE") = rUsrDate
            rsJOB("JH_USRLABEL2") = rUsrLabel2
            rsJOB("JH_USRCHECK2") = rUsrCheck2
            rsJOB("JH_USREDATE2") = rUsrDate2
            rsJOB("JH_USRLABEL3") = rUsrLabel3
            rsJOB("JH_USRCHECK3") = rUsrCheck3
            rsJOB("JH_USREDATE3") = rUsrDate3
            rsJOB("JH_DIV") = rDiv
            rsJOB("JH_DEPTNO") = rDeptNo
            rsJOB("JH_EMP") = rEmp
            rsJOB("JH_GLNO") = rGLNo
            rsJOB("JH_SECTION") = rSect
            rsJOB("JH_REGION") = rRegion
            rsJOB("JH_POSITION_CONTROL") = rPosCtrl
            rsJOB("JH_PAYROLL_ID") = rPayrollID
            rsJOB("JH_GRID") = rGrid
            rsJOB("JH_PAYROLL_CATEGORY") = rPayCateg
            rsJOB("JH_BILLINGRATE") = rBillRate
            'rsJOB("JH_ENDDATE") = rEndDate
            'rsJOB("JH_ENDREAS") = rEndReason
            
            rsJOB("JH_LDATE") = Format(Now, "Short Date")
            rsJOB("JH_LTIME") = Time$
            rsJOB("JH_LUSER") = "999999999"
            rsJOB.Update
            
            rJobID = rsJOB("JH_ID")
            
            rsJOB.Close
            
            'Create a new Salary record as well with same salary
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & Data1.Recordset!ED_EMPNBR & " AND SH_JOB ='" & rJob & "' AND SH_CURRENT<>0 AND SH_SDATE = " & Date_SQL(rSDate)
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsSal.EOF Then
                'Retrieve Data
                rSal = rsSal("SH_SALARY")
                rSalCD = rsSal("SH_SALCD")
                rGrade = rsSal("SH_GRADE")
                rPayP = rsSal("SH_PAYP")
                rReas1 = rsSal("SH_SREAS1")
                rSalPC1 = rsSal("SH_SALPC1")
                rSalChg1 = rsSal("SH_SALCHG1")
                rReas2 = rsSal("SH_SREAS2")
                rSalPC2 = rsSal("SH_SALPC2")
                rSalChg2 = rsSal("SH_SALCHG2")
                rReas3 = rsSal("SH_SREAS3")
                rSalPC3 = rsSal("SH_SALPC3")
                rSalChg3 = rsSal("SH_SALCHG3")
                rCompa = rsSal("SH_COMPA")
                rNextDate = rsSal("SH_NEXTDAT")
                rSComment = rsSal("SH_COMMENT")
                rSComment2 = rsSal("SH_COMMENT2")
                
                'Remove the Current check from the existing current salary record
                rsSal("SH_CURRENT") = False
                rsSal("SH_LDATE") = Format(Now, "Short Date")
                rsSal("SH_LTIME") = Time$
                rsSal("SH_LUSER") = "999999999"
                rsSal.Update
                
                'Add a new Salary records with RTL reason for change
                rsSal.AddNew
                rsSal("SH_COMPNO") = "001"
                rsSal("SH_EMPNBR") = Data1.Recordset!ED_EMPNBR
                rsSal("SH_CURRENT") = True
                rsSal("SH_SDATE") = dlpEDate(0).Text
                rsSal("SH_EDATE") = dlpEDate(0).Text
                rsSal("SH_TRANSDATE") = Format(Now, "SHORT DATE")
                rsSal("SH_SALARY") = rSal
                rsSal("SH_SALCD") = rSalCD
                rsSal("SH_WHRS") = rWHRS
                rsSal("SH_GRADE") = rGrade
                rsSal("SH_PAYP") = rPayP
                rsSal("SH_SREAS1") = "RTL" 'Returns from Temporary Layoff
                rsSal("SH_SALPC1") = 0
                rsSal("SH_SALCHG1") = 0
                rsSal("SH_SREAS2") = rReas2
                rsSal("SH_SALPC2") = 0
                rsSal("SH_SALCHG2") = 0
                rsSal("SH_SREAS3") = rReas3
                rsSal("SH_SALPC3") = 0
                rsSal("SH_SALCHG3") = 0
                rsSal("SH_COMPA") = rCompa
                rsSal("SH_NEXTDAT") = rNextDate
                rsSal("SH_JOB") = rJob
                rsSal("SH_JOB_ID") = rJobID
                rsSal("SH_COMMENT") = rSComment
                rsSal("SH_COMMENT2") = rSComment2
                
                rsSal("SH_LDATE") = Date
                rsSal("SH_LTIME") = Time$
                rsSal("SH_LUSER") = "999999999"
                rsSal("SH_PAYROLL_ID") = rPayrollID
                rsSal.Update
            End If
            
            'Search if any document attached - change reference to new position record
            SQLQ = "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_EMPNBR = " & Data1.Recordset!ED_EMPNBR & " AND DJ_JOB = '" & rJob & "' AND DJ_SDATE = " & Date_SQL(rSDate)
            rsJobDoc.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
            If Not rsJobDoc.EOF Then
                rsJobDoc("DJ_SDATE") = dlpEDate(0).Text
                rsJobDoc("DJ_LUSER") = "999999999"
                rsJobDoc("DJ_LDATE") = Date
                rsJobDoc("DJ_LTIME") = Time$
                rsJobDoc.Update
            End If
            rsJobDoc.Close
            
            MsgBox "A new Position & Salary record has been added with 'RTL' (Returns from Temporary Layoff) Reason for Change." & vbCrLf & "Please check the employee's Position & Salary screen to verify the data.", vbOKOnly, "Employee Position and Salary"
        End If
    End If
End If

'Send Email
If gsEMAIL_ONLEAVECHANGES Then
    MailBody = ""
    If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
        MailBody = GetEmailBodyForSamuel(glbLEE_ID)
        If glbTLAY = "Extending" Then
            MailBody = MailBody & "Leave of Absence Return Date has changed." & vbCrLf & vbCrLf
        ElseIf glbTLAY = "Re-activate" Then
            MailBody = MailBody & "has been Re-activated from Leave of Absence." & vbCrLf & vbCrLf
        End If
        If glbTLAY = "Extending" Then
            MailBody = MailBody & "Change To: " & Data1.Recordset(fdTdate) & vbCrLf
        ElseIf glbTLAY = "Re-activate" Then
            MailBody = MailBody & "Pay Status: " & GetTABLDesc("EDEM", clpCode(1)) & vbCrLf
            MailBody = MailBody & "Pay Status Change Effective (From): " & dlpEDate(0) & vbCrLf
        End If
        If Len(MailBody) > 0 Then
           Screen.MousePointer = DEFAULT
           Call EmailSendingForSamuel
        End If
    Else
        If glbTLAY = "Extending" Then
            MailBody = "The employee's Leave of Absence Return Date has changed." & vbCrLf & vbCrLf
        ElseIf glbTLAY = "Re-activate" Then
            MailBody = "The employee has been Re-activated from Leave of Absence." & vbCrLf & vbCrLf
        End If
        MailBody = MailBody & "Employee #: " & glbLEE_ID & vbCrLf
        'Ticket #22456
        'MailBody = MailBody & "Name: " & RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName) & vbCrLf
        MailBody = MailBody & "Name: " & RTrim$(Data1.Recordset("ED_SURNAME")) & ", " & RTrim$(Data1.Recordset("ED_FNAME")) & vbCrLf
        
        If glbTLAY = "Extending" Then
            MailBody = MailBody & "Changed To: " & Data1.Recordset(fdTdate) & vbCrLf
        ElseIf glbTLAY = "Re-activate" Then
            MailBody = MailBody & "Reason: " & GetTABLDesc("EDEM", clpCode(1)) & vbCrLf
            MailBody = MailBody & "Effective As Of: " & dlpEDate(0) & vbCrLf
        End If
        If Len(MailBody) > 0 Then
           Screen.MousePointer = DEFAULT
           Call imgEmail_Click
        End If
    End If


End If

If Options(1) Then 'Re-activate
    'Ticket #18368
    If glbCompSerial = "S/N - 2259W" Or glbCompSerial = "S/N - 2241W" Then
        Call Employee_Master_Integration(glbLEE_ID)
    End If
    'Ticket #19071 GP Frontenac
    If glbCompSerial = "S/N - 2410W" Then
        Call Employee_Master_Integration(glbLEE_ID, , , , "RetLOA")
    End If
    If glbSamuel Then 'Ticket #20885 Franks 11/18/2011
        Screen.MousePointer = DEFAULT
        Call CheckReptAuth
    End If
    If glbWFC Then 'Ticket #25116 Franks 02/25/2014
        If glbAdv Then
            Call Employee_Master_Integration(glbLEE_ID)
        End If
    End If
    
    Unload Me 'Ticket #24061 Franks 07/16/2013
End If

Screen.MousePointer = DEFAULT

MDIMain.panHelp(0).FloodPercent = 100

Call SET_UP_MODE

'cmdClose.SetFocus
MDIMain.panHelp(0).FloodType = 0

End Sub

'Private Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, X%
Dim xEmplist
'cmdPrint.Enabled = False
Dim rsTB As New ADODB.Recordset
Set rsTB = Data1.Recordset.Clone
xEmplist = ""

Do Until rsTB.EOF
    xEmplist = xEmplist & rsTB("ED_EMPNBR") & ","
    rsTB.MoveNext
Loop
xEmplist = xEmplist & "0"
If glbLinamar Then
    Me.vbxCrystal.WindowTitle = "Temporary Lay Off"
Else
    Me.vbxCrystal.WindowTitle = "Leaves Of Absence"
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    For X% = 0 To 2
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.Formulas(0) = "lblTitle='" & Me.vbxCrystal.WindowTitle & "'"
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGTLAY.rpt"

'Hemu - 10/07/2003 Begin
If glbLinamar Then
    Me.vbxCrystal.Formulas(0) = "lblTitle='Temporary Lay-Off Employees'"
Else
    Me.vbxCrystal.Formulas(0) = "lblTitle='Leave of Absence'"
End If
'Hemu - 10/07/2003 End

Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR} IN [" & xEmplist & "]"
Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
Me.vbxCrystal.Action = 1

'cmdPrint.Enabled = True

End Sub


Sub cmdView_Click()
Dim RHeading As String, xReport, X%
Dim xEmplist
'cmdPrint.Enabled = False
Dim rsTB As New ADODB.Recordset
Set rsTB = Data1.Recordset.Clone
xEmplist = ""
vbxCrystal.Reset
Do Until rsTB.EOF
    xEmplist = xEmplist & rsTB("ED_EMPNBR") & ","
    rsTB.MoveNext
Loop
xEmplist = xEmplist & "0"
If glbLinamar Then
    Me.vbxCrystal.WindowTitle = "Temporary Lay Off"
Else
    Me.vbxCrystal.WindowTitle = "Leaves Of Absence"
End If
If glbSQL Or glbOracle Then
    Me.vbxCrystal.Connect = RptODBC_SQL
Else
    Me.vbxCrystal.Connect = "pwd=petman;"
    For X% = 0 To 2
        Me.vbxCrystal.DataFiles(X%) = glbIHRDB
    Next
End If
Me.vbxCrystal.Formulas(0) = "lblTitle='" & Me.vbxCrystal.WindowTitle & "'"
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RGTLAY.rpt"
Me.vbxCrystal.SelectionFormula = "{HREMP.ED_EMPNBR} IN [" & xEmplist & "]"
Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True

End Sub

'Private Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRJOBEVL", "SELECT")

End Sub

Private Sub dlpTLAYDate_Change()
If IsDate(dlpTLAYDate) And IsDate(Data1.Recordset(fdFdate)) Then
    lblWeeks = Format((CVDate(dlpTLAYDate) - Data1.Recordset(fdFdate)) / 7, "###.0") & " Weeks"
Else
    lblWeeks = ""
End If
If glbWFC Then 'Ticket #19266 Franks 11/29/10
    'If dlpDOther2.Visible Then
    '    dlpDOther2.Text = dlpTLAYDate.Text
    'End If
End If
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMTLAY"
fglbNew = False
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim X%

glbOnTop = "FRMTLAY"

If glbLinamar Then
    Me.Caption = "Temporary Lay Off"
Else
    Me.Caption = "Leaves Of Absence"
    If glbTLAY = "Extending" Then
        Me.Caption = "LOA Date Change"
    End If
    If glbTLAY = "Re-activate" Then
        Me.Caption = "Re-Active from a Leave"
    End If
End If

EESNameSort = True

Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB

If glbLinamar Then
    fdFdate = "ED_USRDAT1"
    fdTdate = "ED_UNION"
    clpCode(1).Text = "ACTI"
    clpCode(3).Text = "INLA"
    clpCode(3).Enabled = False
    Options(3).Visible = True
Else
    fdFdate = "ED_SFDATE"
    fdTdate = "ED_STDATE"
    fraOptions.Height = 1200
    
    'Ticket #29985 - County of Essex - Enter the Union Date
    If glbCompSerial = "S/N - 2192W" Then
        dlpUnion.Left = dlpEDate(1).Left
        lblTitle(1).Visible = True
        lblTitle(1).FontBold = True
        lblTitle(1).Caption = lStr("Union Date")
        dlpUnion.Visible = True
    End If
End If
vbxTrueGrid.Columns(3).DataField = fdFdate
vbxTrueGrid.Columns(4).DataField = fdTdate

Call EEList

If glbAdv Then 'Ticket #14739
    chkATPaidHours.Visible = True
End If

If glbLinamar Then
    If Not gSec_Upd_Terminations Then
    '    cmdOK.Enabled = False
        dlpTLAYDate.Enabled = False
        dlpTermDate.Enabled = False
        dlpEDate(0).Enabled = False
        dlpEDate(1).Enabled = False
        clpCode(1).Enabled = False
        clpCode(3).Enabled = False
    End If
Else
    If Not gSec_Upd_EnterLeave Then
    '    cmdOK.Enabled = False
        dlpTLAYDate.Enabled = False
        dlpTermDate.Enabled = False
        dlpEDate(0).Enabled = False
        dlpEDate(1).Enabled = False
        clpCode(1).Enabled = False
        clpCode(3).Enabled = False
    End If
End If

If vbxTrueGrid.Visible Then Me.vbxTrueGrid.SetFocus

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call INI_Controls(Me)

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
Set frmTLAY = Nothing

End Sub


Private Sub Options_Click(Index As Integer)
fraExtending.Visible = False
fraReActivate.Visible = False
fraTerminate.Visible = False

If Index = 0 Then
    fraExtending.Visible = True
    fraExtending.Top = fraOptions.Top + fraOptions.Height + 500
End If
If Index = 1 Then
    fraReActivate.Visible = True
    If glbWFC Then 'Ticket #23920 Franks 07/04/2013
        fraReActivate.Top = fraOptions.Top
    Else
        fraReActivate.Top = fraOptions.Top + fraOptions.Height + 500
    End If
    fraReActivate.Width = 9075
    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20648 Franks 09/23/2011
        fraReActivate.Height = 1695
    Else
        fraReActivate.Height = 1275
    End If
End If
If Index = 2 Then
    fraTerminate.Visible = True
    fraTerminate.Top = fraOptions.Top + fraOptions.Height + 500
    fraTerminate.Caption = "Terminamte (Will Go To Terminamtion Screen)"
     clpCode(3).Text = "INLA"
    lblTitle(4).Caption = "Termination Reason"
    lblTitle(2).Caption = "Termination Date"
End If
If Index = 3 Then
    fraTerminate.Visible = True
    fraTerminate.Top = fraOptions.Top + fraOptions.Height + 500
    fraTerminate.Caption = "Transfer Out (Will Go To Transfer Out Screen)"
     clpCode(3).Text = "TOUT"
    lblTitle(4).Caption = "Transfer Out Reason"
    lblTitle(2).Caption = "Transfer Out Date"
End If

End Sub

Private Sub Options_GotFocus(Index As Integer)
    Call SetPanHelp(Me.ActiveControl)
End Sub
'Private Sub txtEDate_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtEDate_DblClick(Index As Integer)
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtEDate_GotFocus(Index As Integer)
'    Call SetPanHelp(Me.ActiveControl)
'End Sub
'Private Sub txtEDate_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtEESearch_GotFocus()
    Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub txtTermDate_Change()

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
        
        SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
        If glbOracle Then
            SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
        ElseIf glbLinamar Then
            SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
            SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
        Else
            'Ticket #19871 Franks 02/16/2011
            'SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
            SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
        End If
        SQLQ = SQLQ & "ED_EMPNBR,ED_USRDAT1,ED_UNION,ED_EMP,ED_SFDATE,ED_STDATE,"
        SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER "
        SQLQ = SQLQ & " From HREMP  "
        SQLQ = SQLQ & "Where " & glbSeleDeptUn
        If glbLinamar Then
            SQLQ = SQLQ & " AND ED_EMP='TEMP'"
            If glbTLAY = "Follow-Up" Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT EF_EMPNBR FROM HR_FOLLOW_UP WHERE EF_FREAS='TLAY' AND EF_FDATE<= " & Date_SQL(Date) & ")"
            If glbTLAY = "Extending" Then SQLQ = SQLQ & " AND (DATEDIFF(""ww"",ED_USRDAT1,ED_UNION)<35 or ED_UNION IS NULL)"
        Else
            SQLQ = SQLQ & " AND ED_EMP in (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3<>0)"
        End If
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

Private Function chkTLAY()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset
Dim DgDef As Variant, Response%, Title$  'Ticket #24061 Franks 07/16/2013
Dim rsTemp As New ADODB.Recordset 'Ticket #24094 Franks 07/22/2013

chkTLAY = False
On Error GoTo chkTLAY_Err
If Options(0) Then
    If Not IsDate(dlpTLAYDate.Text) Then
        MsgBox "Extending Date must be valid"
        dlpTLAYDate.SetFocus
        Exit Function
    End If
    If IsDate(dlpTLAYDate.Text) Then
        If glbLinamar Then
            Dim rsTB As New ADODB.Recordset
            Dim qSQLQ
            qSQLQ = "SELECT * FROM HRSTATUS "
            qSQLQ = qSQLQ & " WHERE SC_EMPNBR=" & Data1.Recordset!ED_EMPNBR
            rsTB.Open qSQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            
            If Not rsTB.EOF Then
                If DaysBetween(rsTB("SC_FDATE"), dlpTLAYDate) < 1 Then
                    MsgBox "Date cannot be prior to Leave From Date"
                    dlpTLAYDate.SetFocus
                    Exit Function
                End If
            End If
        Else
            If IsDate(Data1.Recordset("ED_SFDATE")) Then
                If DaysBetween(Data1.Recordset("ED_SFDATE"), dlpTLAYDate) < 1 Then
                    MsgBox "Date cannot be prior to or equal to Leave From Date"
                    dlpTLAYDate.SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
End If


If Options(1) Then
    If Len(clpCode(1).Text) = 0 Then
        MsgBox "Employment Status code is required field"
        clpCode(1).SetFocus
        Exit Function
    Else
        If clpCode(1).Caption = "Unassigned" Then
            MsgBox "Employment Status code must be valid"
            clpCode(1).SetFocus
            Exit Function
        Else
            If chkLeave Then
                MsgBox "Do not use the code for Leave of Absence"
                clpCode(1).SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Not IsDate(dlpEDate(0).Text) Then
        If glbLinamar Then
            MsgBox "Effective From Date must be valid"
        Else
            MsgBox "Effective Date As Of must be valid"
        End If
        dlpEDate(0).SetFocus
        Exit Function
    End If
    If IsDate(dlpEDate(0).Text) And IsDate(dlpEDate(1).Text) Then
        If DaysBetween(dlpEDate(1).Text, dlpEDate(0).Text) < 1 Then
            MsgBox "To Date can not be prior to From Date"
            dlpEDate(0).SetFocus
            Exit Function
        End If
    End If
    If IsDate(dlpEDate(0).Text) And IsDate(Data1.Recordset(fdFdate)) Then
        If DaysBetween(Data1.Recordset(fdFdate), dlpEDate(0).Text) < 0 Then
            MsgBox "As of Date can not be prior to From Date "
            dlpEDate(0).SetFocus
            Exit Function
        End If
    End If
        
    If glbWFC Then 'Ticket #19266 Franks 11/30/2010
        If dlpDOther1.Visible Then
            If Not IsDate(dlpDOther1.Text) Then
                MsgBox lStr("Other Date 1") & " is required field"
                dlpDOther1.SetFocus
                Exit Function
            End If
        End If
    End If
    If glbWFC And frmWFCBenList.Visible Then 'Ticket #24061 Franks 07/16/2013
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "' AND NOT (BM_ENDDATE IS NULL) "
        SQLQ = SQLQ & "AND BM_PCC = 1 " 'NEW - Company % only
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then 'End Date entered, then pop up this message
            Msg$ = "There are Benefit End Dates associated with this employee. Are these benefits re-instated when the employee returns from a leave?"
            Title$ = "Enter Leave Employee"
            DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
            Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
            If Response% = IDNO Then    ' Evaluate response
            '    Exit Function
            Else 'Yes
                '". If "Yes", the user must remove the end dates.
                Exit Function
            End If
        End If
        rsTemp.Close
    End If

    If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20648 Franks 09/23/2011
        If Len(comESalInc.Text) = 0 Then
            MsgBox "Eligible for Salary Increase is required."
            comESalInc.SetFocus
            Exit Function
        End If
    End If
    
    'Ticket #29985 - County of Essex - Enter the Union Date - Mandatory
    If glbCompSerial = "S/N - 2192W" Then
        If Len(Trim(dlpUnion.Text)) = 0 Then
            MsgBox lStr("Union Date") & " is a required field"
            dlpUnion.SetFocus
            Exit Function
        End If
        If Not IsDate(dlpUnion.Text) Then
            MsgBox "Invalid " & lStr("Union Date")
            dlpUnion.SetFocus
            Exit Function
        End If
    End If
End If

If Len(dlpEDate(0)) > 0 And Len(dlpEDate(1)) > 0 Then
    If DaysBetween(dlpEDate(0), dlpEDate(1)) < 0 Then                       'Serbo
        MsgBox "To Date can't be prior to From Date!"                       '
        Me.dlpEDate(0).SetFocus                                             '
        Exit Function                                                       '
    End If
End If

If Options(2) Then
    If Not IsDate(dlpTermDate) Then
        MsgBox "Termination Date must be valid"
        dlpTermDate.SetFocus
        Exit Function
    End If
    If Not glbLinamar Then
        If Len(clpCode(3).Text) = 0 Then
            MsgBox "Termination Reason code is required field"
             clpCode(3).SetFocus
            Exit Function
        Else
            If clpCode(3).Caption = "Unassigned" Then
                MsgBox "Termination Reason  code must be valid"
                 clpCode(3).SetFocus
                Exit Function
            End If
        End If
    End If
End If
If Options(2) Then
    If Not IsDate(dlpTermDate) Then
        MsgBox "Transfer Out Date must be valid"
        dlpTermDate.SetFocus
        Exit Function
    End If
    If Not glbLinamar Then
        If Len(clpCode(3).Text) = 0 Then
            MsgBox "Transfer Out Reason code is required field"
             clpCode(3).SetFocus
            Exit Function
        Else
            If clpCode(3).Caption = "Unassigned" Then
                MsgBox "Transfer Out Reason code must be valid"
                 clpCode(3).SetFocus
                Exit Function
            End If
        End If
    End If
End If
chkTLAY = True

Exit Function

chkTLAY_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkTLAY", "TLAY", "Cancel")
Resume Next

End Function


Private Function EEList()
Dim SQLQ As String, Q As QueryDef
'Dim db As Database
Dim countr   As Integer  ' EEList_Snap is definded at form level

EEList = False         ' if not found - no depts

SQLQ = "SELECT ED_SURNAME,ED_FNAME,"
If glbOracle Then
    SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
ElseIf glbLinamar Then
    SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
    SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR,"
Else
    'Ticket #19871 Franks 02/16/2011
    'SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR,"
    SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR,"
End If
SQLQ = SQLQ & "ED_EMPNBR,ED_USRDAT1,ED_UNION,ED_EMP,ED_SFDATE,ED_STDATE,"
SQLQ = SQLQ & "ED_LDATE,ED_LTIME,ED_LUSER,ED_LTHIRE,ED_SECTION,ED_LDAY  "
SQLQ = SQLQ & " From HREMP  "
SQLQ = SQLQ & "Where " & glbSeleDeptUn
If glbLinamar Then
    SQLQ = SQLQ & " AND ED_EMP='TEMP'"
    If glbTLAY = "Follow-Up" Then SQLQ = SQLQ & " AND ED_EMPNBR IN (SELECT EF_EMPNBR FROM HR_FOLLOW_UP WHERE EF_FREAS='TLAY' AND EF_FDATE<= " & Date_SQL(Date) & ")"
    If glbTLAY = "Extending" Then SQLQ = SQLQ & " AND (DATEDIFF(""ww"",ED_USRDAT1,ED_UNION)<35 or ED_UNION IS NULL)"
Else
    SQLQ = SQLQ & " AND ED_EMP in (SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDEM' AND TB_USR3<>0)"
End If

If EESNameSort = True Then
    If glbOracle Then
        SQLQ = SQLQ & " ORDER BY UPPER(ED_SURNAME), UPPER(ED_FNAME) "
    Else
        SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME "
    End If
Else
    SQLQ = SQLQ & " ORDER BY " & IIf(glbLinamar, "EMPNBR", "ED_EMPNBR")
End If

Data1.RecordSource = SQLQ
Data1.Refresh

EEList = True
Exit Function

EEList_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "VacList", "HREMP", "Select")
Call RollBack '28July99 js

End Function


Private Sub EMPCode_Desc()
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
On Error GoTo EMPCode_Desc_Err
chkLeave.Value = 0

If Len(clpCode(1).Text) > 0 Then
    SQLQ = "SELECT TB_USR3 FROM HRTABL WHERE TB_NAME='EDEM' AND TB_KEY = '" & clpCode(1).Text & "'"
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsTA.EOF Then
        chkLeave.Value = IIf(rsTA("TB_USR3"), 1, 0)
    End If
End If

Exit Sub
EMPCode_Desc_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EMP Code Snap", "TABL", "SELECT")
Call RollBack '29July99 js

End Sub

Private Sub updFollow()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset

On Error GoTo CrFollow_Err

SQLQ = "SELECT * FROM HR_FOLLOW_UP "
SQLQ = SQLQ & " WHERE EF_EMPNBR=" & Data1.Recordset!ED_EMPNBR
If glbLinamar Then
    SQLQ = SQLQ & " AND EF_FREAS='TLAY' "
Else
    SQLQ = SQLQ & " AND EF_FREAS='LOA' "
End If
SQLQ = SQLQ & " AND EF_COMPLETED=0 "
If IsDate(Data1.Recordset(fdTdate)) Then SQLQ = SQLQ & " AND EF_FDATE=" & Date_SQL(Data1.Recordset(fdTdate))
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If Not rsTB.EOF Then
    fglbFollowID = rsTB!EF_FOLLOWUP_ID
Else
    rsTB.AddNew
    Msg = "A Follow Up Record was created!"
    rsTB("EF_COMPNO") = "001"
    rsTB("EF_EMPNBR") = Data1.Recordset("ED_EMPNBR")
    If glbLinamar Then
        rsTB("EF_FREAS") = "TLAY"
    Else
        rsTB("EF_FREAS") = "LOA"
    End If
    rsTB("EF_COMPLETED") = 0
    ' danielk - 02/06/2003 - fixed so it doesn't crash if there's no to/from date in hremp
    If IsDate(Data1.Recordset(fdTdate)) Then
        rsTB("EF_FDATE") = Data1.Recordset(fdTdate)
    Else
        If IsDate(dlpEDate(0).Text) Then
            rsTB("EF_FDATE") = dlpEDate(0).Text
        End If
    End If
End If
If Options(0) Then
    rsTB("EF_FDATE") = CVDate(dlpTLAYDate.Text)
    rsTB("EF_COMMENTS") = rsTB("EF_COMMENTS") & Chr(13) & Chr(10) & "This employee's " & IIf(glbLinamar, "temporary lay-off", "absence leave") & " was extended on " & Format(Date, "MMM DD, YYYY")
    If Msg = "" Then Msg = "A Follow Up Record was updated!"
Else
    rsTB("EF_COMPLETED") = True
    If Options(1) Then
        rsTB("EF_COMMENTS") = rsTB("EF_COMMENTS") & Chr(13) & Chr(10) & "This employee was re-activated on " & Format(dlpEDate(0).Text, "MMM DD, YYYY")
    ElseIf Options(2) Then
        rsTB("EF_COMMENTS") = rsTB("EF_COMMENTS") & Chr(13) & Chr(10) & "This employee was terminated on " & Format(dlpTermDate, "MMM DD, YYYY")
    ElseIf Options(3) Then
        rsTB("EF_COMMENTS") = rsTB("EF_COMMENTS") & Chr(13) & Chr(10) & "This employee was transfer out on " & Format(dlpTermDate, "MMM DD, YYYY")
    End If
    Msg = "A Follow Up Record was marked Complete!"
End If
rsTB("EF_LDATE") = Date
rsTB("EF_LTIME") = Time$
rsTB("EF_LUSER") = glbUserID
rsTB.Update
fglbFollowID = rsTB!EF_FOLLOWUP_ID
rsTB.Close
MsgBox Msg
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub

Private Sub updStatus()   'Laura on 11/2/97
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim xType

On Error GoTo CrFollow_Err

SQLQ = "SELECT * FROM HRSTATUS "
SQLQ = SQLQ & " WHERE SC_EMPNBR=" & Data1.Recordset!ED_EMPNBR
If Options(0) Then
    If IsDate(Data1.Recordset(fdFdate)) Then SQLQ = SQLQ & " AND SC_FDATE=" & Date_SQL(Data1.Recordset(fdFdate))
    If IsDate(Data1.Recordset(fdTdate)) Then SQLQ = SQLQ & " AND SC_TDATE=" & Date_SQL(Data1.Recordset(fdTdate))
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
fAttCode = ""
If Not rsTB.EOF Then
    If Not IsNull(rsTB("SC_ATTREASON")) Then fAttCode = rsTB("SC_ATTREASON")
End If

rsTB.AddNew
rsTB("SC_COMPNO") = "001"
rsTB("SC_EMPNBR") = Data1.Recordset!ED_EMPNBR
rsTB("SC_EMP_TABL") = "EDEM"
rsTB("SC_REASON_TABL") = "SCRE"
If Options(0) Then
    If IsDate(Data1.Recordset(fdFdate)) Then rsTB("SC_FDATE") = Data1.Recordset(fdFdate)
    If IsDate(dlpTLAYDate.Text) Then rsTB("SC_TDATE") = dlpTLAYDate.Text
    rsTB("SC_OLDEMP") = Data1.Recordset!ED_EMP
    rsTB("SC_NEWEMP") = Data1.Recordset!ED_EMP
    rsTB("SC_ATTREASON") = fAttCode
    rsTB("SC_REASON") = "EXTE"
'    rsTB("SC_TYPE") = "HR"
End If
If Options(1) Then
    If IsDate(dlpEDate(0).Text) Then rsTB("SC_FDATE") = dlpEDate(0).Text
    If IsDate(dlpEDate(1).Text) Then rsTB("SC_TDATE") = dlpEDate(1).Text
    rsTB("SC_OLDEMP") = Data1.Recordset!ED_EMP
    rsTB("SC_NEWEMP") = clpCode(1).Text
    rsTB("SC_REASON") = "REAC"
End If
If Options(2) Then
    If IsDate(dlpTermDate) Then rsTB("SC_FDATE") = dlpTermDate
    rsTB("SC_OLDEMP") = Data1.Recordset!ED_EMP
    rsTB("SC_OLDEMP") = Data1.Recordset!ED_EMP
    rsTB("SC_REASON") = "TERM"
End If
If Options(3) Then
    If IsDate(dlpTermDate) Then rsTB("SC_FDATE") = dlpTermDate
    rsTB("SC_OLDEMP") = Data1.Recordset!ED_EMP
    rsTB("SC_OLDEMP") = Data1.Recordset!ED_EMP
    rsTB("SC_REASON") = "TOUT"
End If
rsTB("SC_FOLLOWID") = fglbFollowID
rsTB("SC_JOB") = ReadJob
rsTB("SC_LDATE") = Date
rsTB("SC_LTIME") = Time$
rsTB("SC_LUSER") = glbUserID
rsTB.Update
rsTB.Close
 
'Ticket #16456 update SC_TDATE with Effective Date As for this record
If glbLinamar Then
    If Options(1) Then
        'Ticket #24876 - Delete the Attendance if returning early first because once the SC_TDATE changes below to
        'new return date and the deletion of attendance records becomes an issue.
        SQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_DOA >= " & Date_SQL(dlpEDate(0))
        SQLQ = SQLQ & " AND AD_EMPNBR = " & Data1.Recordset!ED_EMPNBR
        SQLQ = SQLQ & " AND AD_REASON in (SELECT SC_ATTREASON FROM HRSTATUS "
        SQLQ = SQLQ & " WHERE SC_EMPNBR = " & Data1.Recordset!ED_EMPNBR & " AND SC_REASON = '" & fAttCode & "'" 'LOA' "
        SQLQ = SQLQ & " AND SC_FDATE = " & Date_SQL(Data1.Recordset(fdFdate)) & " AND SC_TDATE = "
        SQLQ = SQLQ & Date_SQL(Data1.Recordset(fdTdate)) & ")"
        gdbAdoIhr001.Execute SQLQ
        
        'Ticket #16456 update SC_TDATE with Effective Date As for this record
        If IsDate(Data1.Recordset(fdFdate)) And IsDate(Data1.Recordset(fdTdate)) Then
            SQLQ = "UPDATE HRSTATUS SET SC_TDATE = " & Date_SQL(dlpEDate(0).Text) & " "
            SQLQ = SQLQ & "WHERE SC_EMPNBR=" & Data1.Recordset!ED_EMPNBR & " "
            SQLQ = SQLQ & "AND SC_FDATE=" & Date_SQL(Data1.Recordset(fdFdate)) & " "
            SQLQ = SQLQ & "AND SC_TDATE=" & Date_SQL(Data1.Recordset(fdTdate)) & " "
            SQLQ = SQLQ & "AND SC_TYPE= 'HR' "
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
End If
 
Exit Sub

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Sub

Private Function AUDITTERM()
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xPT, xDiv, XSNAME, XFNAME, xEmpType
Dim SQLQ

On Error GoTo AUDIT_ERR

AUDITTERM = False

rsTB.Open "SELECT ED_EMPNBR,ED_PT,ED_DIV,ED_SURNAME,ED_FNAME,ED_EMPTYPE,ED_EMP FROM HREMP WHERE ED_EMPNBR=" & Data1.Recordset!ED_EMPNBR, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    'xPT = rsTB("ED_PT")
    'xDiv = rsTB("ED_DIV")
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
    XSNAME = rsTB("ED_SURNAME")
    XFNAME = rsTB("ED_FNAME")
    xEmpType = rsTB("ED_EMPTYPE")
'    OLDEMP = rsTB("ED_EMP")
Else
    xPT = ""
    xDiv = ""
    XSNAME = ""
    XFNAME = ""
    xEmpType = ""
'    OLDEMP = ""
End If

Dim strFields As String
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_PTUPL, AU_DIVUPL, AU_EMPTYPE, AU_EMP, AU_UNION, AU_SFDATE, AU_STDATE, AU_COMPNO, "
strFields = strFields & "AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_SURNAME, AU_FNAME, AU_PAYROLL_ID,AU_LTHIRE,AU_PENSION "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False
'
rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_EMPTYPE") = xEmpType

If Options(1) Then
    rsTA("AU_EMP") = clpCode(1).Text
    If glbLinamar Then
        rsTA("AU_UNION") = dlpEDate(0).Text
    Else
        rsTA("AU_SFDATE") = dlpEDate(0).Text
        'Ticket #18306 remove this from interface for Samuel
        If glbCompSerial = "S/N - 2380W" Or glbCompSerial = "S/N - 2384W" Then  'Or glbCompSerial = "S/N - 2382W" Then
        '2380W Vitalaire Ticket #12616
        '2382W Samuel Ticket #18267
        '2384W Ticket #20105 St. Marys Franks 10/11/2011
            If IsDate(dlpEDate(0).Text) Then
                rsTA("AU_LTHIRE") = dlpEDate(0).Text
            End If
        End If
        If IsDate(dlpEDate(1).Text) Then rsTA("AU_STDATE") = dlpEDate(1).Text
        
        'Ticket #29985 - County of Essex - Enter the Union Date
        If glbCompSerial = "S/N - 2192W" Then
            If IsDate(dlpUnion.Text) Then rsTA("AU_UNION") = dlpUnion.Text Else rsTA("AU_UNION") = Null
        End If
    End If
End If
If Options(0) Then
    If glbLinamar Then
        rsTA("AU_UNION") = dlpTLAYDate.Text
    Else
        rsTA("AU_STDATE") = dlpTLAYDate.Text
    End If
End If
'rsTA("AU_SURNAME") = XSNAME
'rsTA("AU_FNAME") = XFNAME

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = Data1.Recordset!ED_EMPNBR
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"

rsTA("AU_TYPE") = "M"

'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    'Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_WORKCOUNTRY,ED_PENSION  FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
        'Ticket #16749 for US ADP Payforce: B - Active Leave
        If Options(1) Then 'Re-active Leave
            If glbWFC Then
                If Not IsNull(rsEmp("ED_WORKCOUNTRY")) Then
                    If rsEmp("ED_WORKCOUNTRY") = "U.S.A." Then
                        rsTA("AU_TYPE") = "B"
                        'Frank 03/01/10 - Begin
                        'From Jerry&MZ:This removes Status Flag 3 in the Banking screen and sends to payroll.
                        'We may need to update this field with a ~ and use that to know to clear the field
                        If Not IsNull(rsEmp("ED_PENSION")) Then
                            If Len((rsEmp("ED_PENSION"))) > 0 Then
                                rsEmp("ED_PENSION") = Null
                                rsEmp.Update
                                rsTA("AU_PENSION") = "-" 'interface will replace "-" to "~"
                            End If
                        End If
                        'Frank 03/01/10 - end
                    End If
                End If
            End If
        End If
    End If
    rsEmp.Close
'End If
rsTA.Update

AUDITTERM = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js

End Function

Private Function ReadJob()
Dim rsTA As New ADODB.Recordset
Dim IJob
ReadJob = ""

rsTA.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & Data1.Recordset!ED_EMPNBR, gdbAdoIhr001, adOpenKeyset
If rsTA.EOF Then Exit Function
ReadJob = rsTA("JH_JOB")
rsTA.Close

End Function
Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsTB As New ADODB.Recordset
Dim SQLQ

If glbTLAY = "Follow-Up" Then
    fraOptions.Visible = True
    Options(1) = True: Call Options_Click(1)
    Options(0).Visible = True
    If glbLinamar Then
        If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
            If IsDate(Data1.Recordset(fdFdate)) And IsDate(Data1.Recordset(fdTdate)) Then
                If DateDiff("ww", Data1.Recordset(fdFdate), Data1.Recordset(fdTdate)) = 35 Then
                    Options(0).Visible = False
                End If
            End If
        End If
    End If
Else
    fraOptions.Visible = False
    If glbTLAY = "Extending" Then Options(0) = True: Call Options_Click(0)
    If glbTLAY = "Re-activate" Then Options(1) = True: Call Options_Click(1)
End If
If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
'cmdOK.Enabled = False
Else
    If Options(0).Value Or Options(1).Value Then 'Ticket #19266 Franks 11/29/10
        If glbWFC Then
            Call WFCOther2Screen(Data1.Recordset("ED_EMPNBR"))
            If Options(1).Value Then 'Ticket #23920 Franks 07/04/2013
                Call WFCBenListScreen(Data1.Recordset("ED_EMPNBR"))
            End If
        End If
        If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #20648 Franks 09/23/2011
            Call SamuelScreenSetup(Data1.Recordset("ED_EMPNBR"))
        End If
    End If
    Exit Sub
End If

If glbLinamar Then
    If IsDate(Data1.Recordset(fdFdate)) Then dlpTLAYDate.Text = DateAdd("ww", 35, Data1.Recordset(fdFdate))
End If
If IsDate(Data1.Recordset(fdTdate)) Then
    dlpEDate(0).Text = Data1.Recordset(fdTdate)
Else
        dlpEDate(0).Text = Date
End If

If glbLinamar Then
    lblTitle(0) = "Effective Date As Of"
    dlpEDate(1).Visible = False
    lblTitle(1).Visible = False
Else
    lblTitle(0) = "Effective Date From"
    dlpEDate(1).Text = ""
    'Hemu - 18/08/2003 Begin - Jerry asked to hide it
    'dlpEDate(1).Visible = True
    'lblTitle(1).Visible = True
    'Hemu - 18/08/2003 End
End If
If IsDate(Data1.Recordset(fdTdate)) Then
    dlpTermDate = Data1.Recordset(fdTdate)
Else
    dlpTermDate = Date
End If
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
RelateMode = RelateTermEmp
End Property

Public Property Get UpdateRight() As Boolean
If glbLinamar Then
    UpdateRight = gSec_Upd_Terminations
Else
    UpdateRight = gSec_Upd_EnterLeave
End If
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property
Public Property Get Updateble() As Boolean
Updateble = True
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False
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
'ElseIf rsEMP.EOF Then
'    UpdateState = NoRecord
'    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub
Private Sub lblEEID_Change()
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
'    frmEBENEFITS.Caption = "Benefits / Beneficiaries - " & Left$(glbLEE_SName, 5)
 '   frmTLAY.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
'lblEEID = glbLEE_ID
'lblEENum = ShowEmpnbr(lblEEID)
End Sub


Private Function updAttendance()
Dim SQLQ As String
Dim rsJOB As New ADODB.Recordset, rsDup As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsATT As New ADODB.Recordset
Dim xDays
Dim X, xDATE, xDup
Dim WSQLQ, ESQLQ, Result
Dim TSQLQ
Dim Msg$
Dim AskWeekend, SkipWeekend
Dim xWeekDay
Dim xFromDate, xToDate
Dim xHours, xSHIFT, xSuper, xIncID, xSEN, xEMELEA, xINDICATOR
Dim xKey
updAttendance = False
On Error GoTo updAttendance_Err

If Len(fAttCode) = 0 Then Exit Function

If Not IsDate(Data1.Recordset(fdTdate)) Then Exit Function

xFromDate = DateAdd("d", 1, Data1.Recordset(fdTdate))
If Not IsDate(dlpTLAYDate.Text) Then Exit Function
xToDate = dlpTLAYDate.Text

Screen.MousePointer = HOURGLASS

xHours = 0
xSHIFT = Null
xSuper = Null
rsJOB.Open "SELECT JH_DHRS,JH_REPTAU,JH_SHIFT FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenForwardOnly
If Not rsJOB.EOF Then
    If IsNumeric(rsJOB("JH_DHRS")) Then xHours = rsJOB("JH_DHRS") Else xHours = 0
    xSuper = rsJOB("JH_REPTAU")
    xSHIFT = rsJOB("JH_SHIFT")
End If
rsJOB.Close
rsTB.Open "SELECT * FROM HRTABL WHERE TB_NAME='ADRE' AND TB_KEY='" & fAttCode & "'", gdbAdoIhr001, adOpenForwardOnly
xIncID = 0
xSEN = 0
xEMELEA = 0
xINDICATOR = 0
If Not rsTB.EOF Then
    xSEN = rsTB("TB_SEN")
    xEMELEA = rsTB("TB_USR3")
    xINDICATOR = rsTB("TB_INDICATOR")
End If
rsTB.Close

If UCase(fAttCode) = "OT15" Then xHours = xHours * 1.5
If UCase(fAttCode) = "OT20" Then xHours = xHours * 2

'City of Timmins - Ticket #16168
If glbCompSerial = "S/N - 2375W" Then
    If UCase(fAttCode) = "OT05" Then xHours = xHours * 0.5
    If UCase(fAttCode) = "OT25" Then xHours = xHours * 2.5
End If

If Len(xToDate) = 0 Then
    xDays = 0
Else
    xDays = DateDiff("d", xFromDate, xToDate)
End If
xDATE = xFromDate
AskWeekend = True

'Ticket #13183 Frank 06/13/07
'If made the date earlier
'the program should delete the attendance records which are not in this range
If xDays < 0 Then
    TSQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_REASON = '" & fAttCode & "' "
    TSQLQ = TSQLQ & " AND AD_DOA > " & Date_SQL(xToDate)
    TSQLQ = TSQLQ & " AND AD_DOA < " & Date_SQL(xFromDate)
    TSQLQ = TSQLQ & " AND AD_EMPNBR =" & glbLEE_ID
    gdbAdoIhr001.Execute TSQLQ
End If

For X = 0 To xDays
   xWeekDay = Weekday(xDATE)
   If xWeekDay = 7 Or xWeekDay = 1 Then
        If AskWeekend Then
            Msg$ = "Do you want to exclude Saturday/Sunday for Attendance Records?"
            AskWeekend = False
            SkipWeekend = False
            If MsgBox(Msg$, 36) = 6 Then
                SkipWeekend = True
                xDATE = DateAdd("d", IIf(xWeekDay = 7, 2, 1), xDATE)
                X = X + IIf(xWeekDay = 7, 2, 1)
            End If
        Else
            If SkipWeekend Then
                xDATE = DateAdd("d", IIf(xWeekDay = 7, 2, 1), xDATE)
                X = X + IIf(xWeekDay = 7, 2, 1)
            End If
        End If
    End If
    If Len(xToDate) > 0 Then
        If CVDate(xDATE) > CVDate(xToDate) Then Exit For
    Else
        If CVDate(xDATE) > CVDate(xFromDate) Then Exit For
    End If
    
    TSQLQ = "SELECT AD_EMPNBR FROM HR_ATTENDANCE "
    TSQLQ = TSQLQ & " WHERE AD_REASON = '" & fAttCode & "' "
    TSQLQ = TSQLQ & " AND AD_DOA = " & Date_SQL(xDATE)
    TSQLQ = TSQLQ & " AND AD_EMPNBR =" & glbLEE_ID
    rsDup.Open TSQLQ, gdbAdoIhr001, adOpenKeyset
    If Not rsDup.EOF Then
        Msg$ = "Reason: " & fAttCode & Chr(10) & " Date: " & xDATE & Chr(10) & Chr(10)
        Msg$ = Msg$ & rsDup.RecordCount & " duplicates found in Attendance Master. " & Chr(10) & Chr(10)
        Msg$ = Msg$ & "Click Yes to post all Attendance records including duplicates." & Chr(10)
        Msg$ = Msg$ & "Click No to post all non-duplicate Attendance records." & Chr(10)
        Result = MsgBox(Msg$, vbYesNo, "Duplicates Found")
        If Result = vbYes Then
            xDup = True
        Else
            xDup = False
        End If
    Else
        xDup = True
    End If
    
    rsDup.Close
    
    If xDup Then
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=0"
        rsATT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        rsATT.AddNew
        rsATT("AD_EMPNBR") = glbLEE_ID
        rsATT("AD_COMPNO") = "001"
        rsATT("AD_DOA") = xDATE
        rsATT("AD_REASON") = fAttCode
        rsATT("AD_HRS") = xHours
        'rsATT("AD_COMM") = ""
        rsATT("AD_SHIFT") = xSHIFT
        rsATT("AD_SUPER") = xSuper
        rsATT("AD_INCID") = xIncID
        rsATT("AD_SEN") = xSEN
        rsATT("AD_EMELEA") = xEMELEA
        rsATT("AD_INDICATOR") = xINDICATOR
        rsATT("AD_LDATE") = Date
        rsATT("AD_LTIME") = Time$
        rsATT("AD_LUSER") = glbUserID
        rsATT.Update
        If glbAdv Then 'Ticket #14739
            xKey = rsATT("AD_EMPNBR")
            xKey = xKey & "|" & Format(rsATT("AD_DOA"), "dd-mmm-yyyy")
            xKey = xKey & "|" & rsATT("AD_REASON")
            If chkATPaidHours.Value Then
                Call Attendance_Master_Integration(xKey, rsATT("AD_ATT_ID"))
            Else
                Call Attendance_Master_Integration(xKey, rsATT("AD_ATT_ID"), , "YES")
            End If
        End If
        rsATT.Close
    End If
    xDATE = DateAdd("d", 1, xDATE)
Next
Call EntReCalc("ED_EMPNBR=" & glbLEE_ID)
Call EntReCalcHr


updAttendance = True

Exit Function

updAttendance_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "updAttendance", "Attendance", "Insert")
updAttendance = False
Resume Next

End Function

'Ticket #27742 Franks 11/10/2015
'Private Sub CheckHRTABLCode(xName, xKey, xKeyDesc)
'Dim RSTABL As New ADODB.Recordset
'Dim SQLQ
'    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & xName & "' AND TB_KEY = '" & xKey & "' "
'    RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If RSTABL.EOF Then
'        RSTABL.AddNew
'        RSTABL("TB_COMPNO") = "001"
'        RSTABL("TB_NAME") = xName
'        RSTABL("TB_KEY") = xKey
'        RSTABL("TB_DESC") = Trim(xKeyDesc)
'        RSTABL("TB_LDATE") = Format(Now, "Short Date")
'        RSTABL("TB_LTIME") = Time$
'        RSTABL("TB_LUSER") = "999999999"
'        RSTABL.Update
'    End If
'    RSTABL.Close
'End Sub

Sub EmailSendingForSamuel()
Dim xEmail
Dim xToEmail As String
Dim xEmpName As String
Dim xEmailSubject As String, xBranch  As String

On Error GoTo Email_Err
    If gsEMAIL_ONLEAVECHANGES Then

        xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
        End If
         
        If Len(xToEmail) > 0 Then
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            'frmSendEmail.txtSubject.Text = "info:HR Leave Changes Notice"
            'Ticket #18578
            xEmpName = ""
            xBranch = ""
            If Not Data1.Recordset.EOF Then
                xEmpName = " - " & Data1.Recordset("ED_SURNAME") & ", " & Data1.Recordset("ED_FNAME")
                xBranch = Data1.Recordset("ED_SECTION")
            End If
            'Ticket #18755
            'frmSendEmail.txtSubject.Text = "info:HR Leave Changes Notice" & xEmpName 'lblEEName
            If Len(xBranch) > 0 Then
                xBranch = xBranch & " - "
            End If
            xEmailSubject = "info:HR Leave Changes Notice - " & xBranch & xEmpName
            frmSendEmail.txtSubject.Text = xEmailSubject
        
            frmSendEmail.txtBody.Text = MailBody
            'frmSendEmail.Show 1
            MDIMain.panHelp(0).FloodType = 0
            MDIMain.panHelp(0).Caption = "Sending email..."
            frmSendEmail.Tag = ""
            frmSendEmail.cmdSend_Click
            Do
                DoEvents
            Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
            ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
            If frmSendEmail.Tag = "DONE" Then
                Unload frmSendEmail
            Else
                Unload frmSendEmail
            End If
            MDIMain.panHelp(0).Caption = ""
            MDIMain.panHelp(0).FloodType = 1
        End If
        
    End If

Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    'Resume Next
    Exit Sub

End Sub
Public Sub imgEmail_Click()
Dim xEmail
Dim xToEmail As String
Dim xEmpName As String
On Error GoTo Email_Err
    If gsEMAIL_ONLEAVECHANGES Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            
        'Ticket #18235 - Email on Leave Changes
        If glbCompSerial = "S/N - 2382W" Then  'Samuel
            xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            End If
        Else
            'Ticket #20317 - More Emails for everyone
            xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES", glbLEE_ID)
            If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
                xToEmail = GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            End If
        End If
        'Ticket #18235 - End
            
        'If Len(xEmail) > 0 Then    'Hemu - (Ticket #11562) - Jerry asked to remove the check for email address presence.
            frmSendEmail.txtTo.Text = xToEmail 'GetComPreferEmail("EMAIL_ONLEAVECHANGES")
            frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            xEmpName = ""
            'frmSendEmail.txtSubject.Text = "info:HR Leave Changes Notice"
            If Not Data1.Recordset.EOF Then
                xEmpName = " - " & Data1.Recordset("ED_SURNAME") & ", " & Data1.Recordset("ED_FNAME")
            End If
            frmSendEmail.txtSubject.Text = "info:HR Leave Changes Notice" & xEmpName
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        'Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            'End If
        '    MsgBox "There is no email address for the 'Email Notification on Leave Changes' on Company Preference screen. "
        'End If
    End If

Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next

End Sub
Private Sub WFCOther2Screen(xEmpNo)
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
    
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    If Options(0).Value Then 'LOA Date Change
        fraExtending.Height = 735
        lbOtherDate2.Visible = False
        dlpDOther2.Visible = False
        
        SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsEmpee.EOF Then
            Exit Sub
        Else
            If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
            If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
            If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
            If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
        End If
        rsEmpee.Close
        
        'No NGS Sub Group, skip
        If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub
    
        
        xNGSStart = ""
        dlpDOther2.Text = ""
        SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1,ER_OTHERDATE2 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
        rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmpOther.EOF Then
            If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
                xNGSStart = rsEmpOther("ER_OTHERDATE1")
            End If
            If IsDate(rsEmpOther("ER_OTHERDATE2")) Then
                dlpDOther2.Text = rsEmpOther("ER_OTHERDATE2")
            End If
        End If
        rsEmpOther.Close
        'No NGS Effective Date, skip
        If Len(xNGSStart) = 0 Then Exit Sub
        fraExtending.Height = 975
        lbOtherDate2.Caption = lStr("Other Date 2")
        lbOtherDate2.Visible = True
        dlpDOther2.Visible = True
    End If
    
    If Options(1).Value Then 'Re-Activate from a Leave
        lbOtherDate1.Visible = False
        dlpDOther1.Visible = False
        
        SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsEmpee.EOF Then
            Exit Sub
        Else
            If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
            If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
            If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
            If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
        End If
        rsEmpee.Close
        
        'No NGS Sub Group, skip
        If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub
          
        xNGSStart = ""
        xNGStmpDate = ""
        SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1,ER_OTHERDATE2 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
        rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmpOther.EOF Then
            If Not IsNull(rsEmpOther("ER_OTHERDATE1")) Then
                If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
                    xNGSStart = rsEmpOther("ER_OTHERDATE1")
                End If
            End If
            If Not IsNull(rsEmpOther("ER_OTHERDATE2")) Then
                If IsDate(rsEmpOther("ER_OTHERDATE2")) Then
                    xNGStmpDate = rsEmpOther("ER_OTHERDATE2")
                End If
            End If
        End If
        rsEmpOther.Close
        'No NGS Effective Date, skip
        If Len(xNGSStart) = 0 Then Exit Sub
        lbOtherDate1.Top = lblTitle(0).Top
        dlpDOther1.Top = dlpEDate(0).Top
        lbOtherDate1.Caption = lStr("Other Date 1")
        lbOtherDate1.Visible = True
        dlpDOther1.Visible = True
        'Ticket #22109
        dlpDOther1.Text = xNGSStart
    End If
End Sub
Private Sub WFC_NGS_Trans(xEmpNo, SvDOther, xType)
Dim xLDate
    If Not glbNGS_OnFlag Then
        Exit Sub
    End If
    
    If xType = "LOA Change" Then
        If IsDate(SvDOther) Then
            Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE2", CVDate(SvDOther))
        'Else
        '    Call Upt_EmpOtherByField(glbLEE_ID, "ER_OTHERDATE2", Null)
        End If
        If IsDate(SvDOther) Then
            xLDate = SvDOther 'Date
            Call NGSAuditAdd(xEmpNo, "M", "LOA Date Change", lStr("Other Date 2"), "", SvDOther, xLDate)
        End If
    End If
    If xType = "Re-Active" Then
        If IsDate(SvDOther) Then
            Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE1", CVDate(SvDOther))
            Call Upt_EmpOtherByField(xEmpNo, "ER_OTHERDATE2", Null)
            xLDate = SvDOther 'Date
            Call NGSAuditAdd(xEmpNo, "M", "Re-Activate from a Leave", lStr("Other Date 1"), "", SvDOther, xLDate)
            If IsDate(xNGStmpDate) Then
                Call NGSAuditAdd(xEmpNo, "M", "Re-Activate from a Leave", lStr("Other Date 2"), xNGStmpDate, "", xLDate)
            End If
        End If
    End If
End Sub

Private Sub SAMUEL_Trans(xEmpNo, SvDOther, xType)
Dim xLDate
        xLDate = SvDOther
        Call SamuelAuditAdd(xEmpNo, "M", "Re-Activate from a Leave", "Return Leave", "", SvDOther, xLDate)
End Sub

Private Sub SamuelScreenSetup(xEmpNo)
Dim rsEmpee As New ADODB.Recordset
Dim SQLQ As String

comESalInc.Clear
comESalInc.AddItem "Yes"
comESalInc.AddItem "No"
lblTitle(5).Visible = True
comESalInc.Visible = True

glbEmpDiv = ""
glbEmpAdminBy = ""
glbEmpSection = ""
glbEmpRegion = ""
SQLQ = "SELECT ED_EMPNBR, ED_ADMINBY, ED_DIV, ED_SECTION, ED_REGION FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic 'ED_VADIM1
If rsEmpee.EOF Then
    Exit Sub
Else
    If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
    If IsNull(rsEmpee("ED_ADMINBY")) Then glbEmpAdminBy = "" Else glbEmpAdminBy = rsEmpee("ED_ADMINBY")
    If IsNull(rsEmpee("ED_SECTION")) Then glbEmpSection = "" Else glbEmpSection = rsEmpee("ED_SECTION")
    If IsNull(rsEmpee("ED_REGION")) Then glbEmpRegion = "" Else glbEmpRegion = rsEmpee("ED_REGION")
End If
rsEmpee.Close

End Sub

Private Sub CheckReptAuth() 'Ticket #20885 Franks 11/18/2011 for Samuel
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean
Dim xMsg As String
    xFlag1 = False
    'check if this employee is a Reporting Authority
    If IsReportAuth(glbLEE_ID) Then
        xFlag1 = True
    End If

    If xFlag1 Then
        xMsg = "This employee has been assigned as a Reporting Authority on other employee files."
        xMsg = xMsg & "  Will this Return from LOA affect the Reporting Authority structures?"
        frmMsgYesNoUn.lblMsg.Caption = xMsg
        frmMsgYesNoUn.lblMsg.Alignment = 0
        frmMsgYesNoUn.Show 1
        If glbMsgCustomVal = 1 Or glbMsgCustomVal = 3 Then
            'create a report to show the employee list
            Call CreateEmpList4ReportAuth(Data1.Recordset("ED_EMPNBR")) '(glbLEE_ID)
            'show the report - begin
            Me.vbxCrystal.Reset
            Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "RZEmpList2.rpt"
            If Len(glbstrSelCri) >= 0 Then
                Me.vbxCrystal.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
            End If
            'Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & Data1.Recordset("ED_SURNAME") & "," & Data1.Recordset("ED_FNAME") & "'"
            'Ticket #21669 Franks 03/01/2012
            xMsg = Replace(Data1.Recordset("ED_SURNAME") & "," & Data1.Recordset("ED_FNAME"), "'", "''")
            Me.vbxCrystal.Formulas(0) = "rTitle='Employee List for Reporting Authority " & xMsg & "'"
            Me.vbxCrystal.Connect = RptODBC_SQL
            Me.vbxCrystal.WindowTitle = "Employee List for Reporting Authority " & Data1.Recordset("ED_SURNAME") & "," & Data1.Recordset("ED_FNAME")
            Me.vbxCrystal.Destination = 0
            Me.vbxCrystal.Action = 1
            Me.vbxCrystal.Reset
            'show the report - end
        End If
    End If
    
End Sub

Private Sub WFCBenListScreen(xEmpNo) 'Ticket #23920 Franks 07/04/2013
Dim rsLEmp As New ADODB.Recordset
Dim rslBen As New ADODB.Recordset
Dim SQLQ As String
    frmWFCBenList.Visible = False
    frmWFCBenList.Top = fraReActivate.Top + fraReActivate.Height + 100
    frmWFCBenList.Left = fraReActivate.Left
    frmWFCBenList.Width = 10575
    frmWFCBenList.Height = 3135
    chkAllDates.Caption = "All End Dates"
    lblTitle(6).Caption = lStr("Last Day")
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsLEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xoLDAY = ""
    If Not rsLEmp.EOF Then
        If Not IsNull(rsLEmp("ED_WORKCOUNTRY")) Then
            If rsLEmp("ED_WORKCOUNTRY") = "U.S.A." Then
                If Not IsNull(rsLEmp("ED_VADIM1")) Then
                    If Len(rsLEmp("ED_VADIM1")) > 0 Then
                        If Not IsNull(rsLEmp("ED_LDAY")) Then
                            dlpLastDate.Text = rsLEmp("ED_LDAY")
                        Else
                            dlpLastDate.Text = ""
                        End If
                        xoLDAY = dlpLastDate.Text
                        Call UpdateBenefitGroup(xEmpNo)
                        
                        Data2.ConnectionString = glbAdoIHRDBW
                        SQLQ = "SELECT * FROM HRBENGRPLIST "
                        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
                        Data2.RecordSource = SQLQ
                        Data2.Refresh

                        frmWFCBenList.Visible = True
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub UpdateBenefitGroup(xEmpNo) 'Ticket #23920 Franks 07/02/2013
Dim rsBGMST As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim rsBGEE As New ADODB.Recordset
Dim RSTABL As New ADODB.Recordset
Dim SQLQ As String
Dim BelongOldGroup As Boolean
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001W.CommitTrans

    gdbAdoIhr001W.BeginTrans
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "ORDER BY BF_BCODE, BF_EDATE "

    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsBGMST.EOF
        rsBGTMP.AddNew
        rsBGTMP("BM_COMPNO") = "001"
        rsBGTMP("BM_BENEFIT_GROUP") = rsBGMST("BF_GROUP")
        rsBGTMP("BM_BCODE") = rsBGMST("BF_BCODE")
        rsBGTMP("BM_EDATE") = rsBGMST("BF_EDATE")
        rsBGTMP("BM_ENDDATE") = rsBGMST("BF_CEASEDATE") 'New
        rsBGTMP("BM_CHECK") = 1
        rsBGTMP("BM_COVER") = rsBGMST("BF_COVER")
        rsBGTMP("BM_AMT") = rsBGMST("BF_AMT")
        rsBGTMP("BM_PPAMT") = rsBGMST("BF_PPAMT")
        rsBGTMP("BM_UNITCOST") = rsBGMST("BF_UNITCOST")
        rsBGTMP("BM_PCE") = rsBGMST("BF_PCE")
        rsBGTMP("BM_PCC") = rsBGMST("BF_PCC")
        rsBGTMP("BM_ECOST") = rsBGMST("BF_ECOST")
        rsBGTMP("BM_CCOST") = rsBGMST("BF_CCOST")
        rsBGTMP("BM_TCOST") = rsBGMST("BF_TCOST")
        rsBGTMP("BM_MAXDOL") = rsBGMST("BF_MAXDOL")
        rsBGTMP("BM_PREMIUM") = rsBGMST("BF_PREMIUM")
        rsBGTMP("BM_PER") = rsBGMST("BF_PER")
        rsBGTMP("BM_MTHCCOST") = rsBGMST("BF_MTHCCOST")
        rsBGTMP("BM_MTHECOST") = rsBGMST("BF_MTHECOST")
        rsBGTMP("BM_TAXBEN") = rsBGMST("BF_TAXBEN")
        rsBGTMP("BM_SALARYDEPENDANT") = rsBGMST("BF_SALARYDEPENDANT")
        rsBGTMP("BM_MINIMUM") = rsBGMST("BF_MINIMUM")
        rsBGTMP("BM_FACTOR") = rsBGMST("BF_FACTOR")
        rsBGTMP("BM_ROUND") = rsBGMST("BF_ROUND")
        rsBGTMP("BM_MAXIMUM") = rsBGMST("BF_MAXIMUM")
        rsBGTMP("BM_NEXTNEAREST") = rsBGMST("BF_NEXTNEAREST")
        rsBGTMP("BM_TAXAMOUNT") = rsBGMST("BF_TAXAMOUNT")
        rsBGTMP("BM_WAITPERIOD") = rsBGMST("BF_WAITPERIOD")
        rsBGTMP("BM_DWM") = rsBGMST("BF_DWM")
        rsBGTMP("BM_PERORDOLL") = rsBGMST("BF_PERORDOLL")
        rsBGTMP("BM_POLICY") = rsBGMST("BF_POLICY")
        rsBGTMP("BM_RATELEVEL") = rsBGMST("BF_RATELEVEL")
        rsBGTMP("BM_COMMENTS") = rsBGMST("BF_COMMENTS")
        rsBGTMP("BM_PTAX") = rsBGMST("BF_PTAX")
        rsBGTMP("BM_ACTION") = "Add"
        rsBGTMP("BM_WRKEMP") = glbUserID
        
        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BF_BCODE") & "' "
        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not RSTABL.EOF Then
            rsBGTMP("BM_BCODE_DESC") = RSTABL("TB_DESC")
        End If
        RSTABL.Close
        rsBGTMP.Update
        rsBGMST.MoveNext
    Loop
    rsBGTMP.Close
    rsBGMST.Close
    gdbAdoIhr001W.CommitTrans
    Call Pause(1)

End Sub

Private Sub chkAllDates_Click() 'Ticket #23920 Franks 07/04/2013
Dim SQLQ As String
Dim xID As Long
        If Not Data2.Recordset.EOF Then
            xID = Data2.Recordset("BM_BENE_ID")
            If chkAllDates.Value Then 'checked
                SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = Null WHERE BM_WRKEMP = '" & glbUserID & "' "
                SQLQ = SQLQ & "AND BM_PCC = 1 "
                'SQLQ = SQLQ & "AND NOT (BM_BENE_ID = " & xID & ") "
                gdbAdoIhr001.Execute SQLQ
                Data2.Refresh
                SQLQ = "BM_BENE_ID = " & xID
                Data2.Recordset.Find SQLQ
            Else 'unchecked
                'SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = " & Date_SQL(dlpEndDate.Text) & " WHERE BM_WRKEMP = '" & glbUserID & "' "
                'SQLQ = SQLQ & "AND BM_PCC = 1 "
                'gdbAdoIhr001.Execute SQLQ
                'Data2.Refresh
                'SQLQ = "BM_BENE_ID = " & xID
                'Data2.Recordset.Find SQLQ
            End If
        End If
End Sub

Private Sub WFCUpdate_Value() 'Ticket #23920 Franks 07/02/2013
Dim SQLQ As String
Dim xID As Long
If IsDate(dlpEndDate.Text) Or Len(dlpEndDate.Text) = 0 Then
    If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
        xID = Data2.Recordset("BM_BENE_ID")
        If IsDate(dlpEndDate.Text) Then
            If Year(dlpEndDate.Text) > 1900 And Year(dlpEndDate.Text) < 2050 Then
                Data2.Recordset("BM_ENDDATE") = dlpEndDate.Text
            Else
                Data2.Recordset("BM_ENDDATE") = Null
            End If
        Else
            Data2.Recordset("BM_ENDDATE") = Null
        End If
        Data2.Recordset.Update
        Data2.Refresh
        SQLQ = "BM_BENE_ID = " & xID
        Data2.Recordset.Find SQLQ
    End If
End If
End Sub

Private Sub vbxTrueGrid1_BeforeRowColChange(Cancel As Integer)
Call WFCUpdate_Value 'Ticket #23920 Franks 07/04/2013
End Sub

'Ticket #23920 Franks 07/03/2013
Private Sub vbxTrueGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not (Data2.Recordset.EOF Or Data2.Recordset.BOF) Then
        If IsNull(Data2.Recordset("BM_ENDDATE")) Then
            dlpEndDate.Text = ""
        Else
            dlpEndDate.Text = Data2.Recordset("BM_ENDDATE")
        End If
    End If
End Sub

Private Sub WFC_NGSBenEndDateUpt(xEmpNo) 'Ticket #23920 Franks 07/04/2013
Dim SQLQ, xACT
Dim rsBN As New ADODB.Recordset
Dim rsEmpBN As New ADODB.Recordset
Dim xTemp
Dim xDate1, xDate2
    'update Last Day xoLDAY
    If Not xoLDAY = dlpLastDate.Text Then 'changed
        If IsDate(dlpLastDate.Text) Then
            SQLQ = "UPDATE HREMP SET ED_LDAY = " & Date_SQL(dlpLastDate.Text) & " "
            SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo 'Ticket #24588 Franks 11/01/2013
            gdbAdoIhr001.Execute SQLQ
        Else
            SQLQ = "UPDATE HREMP SET ED_LDAY = Null "
            SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo 'Ticket #24588 Franks 11/01/2013
            gdbAdoIhr001.Execute SQLQ
        End If
        If IsDate(dlpLastDate.Text) Then
            Call AUDITBENF(xEmpNo, False, , "Y", dlpLastDate.Text)
        End If
    End If
    
    SQLQ = "SELECT * FROM HRBENGRPLIST "
    SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
    rsBN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsBN.EOF
        SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_BCODE = '" & rsBN("BM_BCODE") & "' "
        If Not IsNull(rsBN("BM_EDATE")) Then SQLQ = SQLQ & "AND BF_EDATE = " & Date_SQL(rsBN("BM_EDATE")) & " "
        If rsEmpBN.State <> 0 Then rsEmpBN.Close
        rsEmpBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpBN.EOF Then
            If IsNull(rsEmpBN("BF_CEASEDATE")) Then xDate1 = CVDate("01/01/1900") Else xDate1 = CVDate(rsEmpBN("BF_CEASEDATE"))
            If IsNull(rsBN("BM_ENDDATE")) Then xDate2 = CVDate("01/01/1900") Else xDate2 = CVDate(rsBN("BM_ENDDATE"))
            rsEmpBN("BF_CEASEDATE") = rsBN("BM_ENDDATE")
            rsEmpBN.Update
            If Not xDate1 = xDate2 Then 'BF_CEASEDATE was changed
                'If xDate2 > CVDate("01/01/1900") Then
                    'update hraudit - begin
                    Call AUDITBENF(xEmpNo, False, rsEmpBN)
                    'update hraudit - end
                'End If
            End If
        End If
        rsEmpBN.Close
        rsBN.MoveNext
    Loop
    rsBN.Close
End Sub

Private Function AUDITBENF(xEmpNo, xlocNewRec As Boolean, Optional rslBen As ADODB.Recordset, Optional xIsWorkDay = "N", Optional xLastDate) 'Ticket #23920 Franks 07/03/2013
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
Dim ACTX
Dim NBCode, NPPAMT, NMTHCOMP, NMTHEMP, NBAMT, NPPE, NPCC, NMAXDOL, NEDate, NCOVER, NTCOST
Dim xTermSEQ
Dim SQLQ As String

'''On Error GoTo AUDIT_ERR
AUDITBENF = False

If xlocNewRec Then
    ACTX = "A"
Else
    ACTX = "M"
End If

xTermSEQ = 0
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
Else
    SQLQ = "SELECT ED_PT,ED_DIV FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS,AU_CEASEDATE,AU_LDAY "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

xADD = False

If xIsWorkDay = "N" Then
    NBCode = ""
    NPPAMT = ""
    NMTHCOMP = ""
    NMTHEMP = ""
    NBAMT = ""
    NPPE = ""
    NPCC = ""
    NMAXDOL = ""
    NEDate = ""
    NCOVER = ""
    NTCOST = ""
    NBCode = rslBen("BF_BCODE")
    If Not IsNull(rslBen("BF_EDATE")) Then NEDate = rslBen("BF_EDATE")
    ''If Not IsNull(rslBen("BF_PPAMT")) Then NPPAMT = rslBen("BF_PPAMT")
    ''If Not IsNull(rslBen("BF_MTHCCOST")) Then NMTHCOMP = rslBen("BF_MTHCCOST")
    ''If Not IsNull(rslBen("BF_MTHECOST")) Then NMTHEMP = rslBen("BF_MTHECOST")
    ''If Not IsNull(rslBen("BF_AMT")) Then NBAMT = rslBen("BF_AMT")
    ''If Not IsNull(rslBen("BF_PCC")) Then NPCC = rslBen("BF_PCC")
    ''If Not IsNull(rslBen("BF_PCE")) Then NPPE = rslBen("BF_PCE")
    ''If Not IsNull(rslBen("BF_MAXDOL")) Then NMAXDOL = rslBen("BF_MAXDOL")
    ''If Not IsNull(rslBen("BF_COVER")) Then NCOVER = rslBen("BF_COVER")
    ''If Not IsNull(rslBen("BF_TCOST")) Then NTCOST = rslBen("BF_TCOST")
    ''
    ''If OBCode <> NBCode Then GoTo MODUPD
    '''If OPPE <> NPPE Or OPCC <> NPCC Then GoTo MODUPD
    ''If OPPAMT <> NPPAMT Or OMAXDOL <> NMAXDOL Then GoTo MODUPD
    '''If OMTHCOMP <> NMTHCOMP Or OMTHEMP <> NMTHEMP Then GoTo MODUPD
    ''If OBAMT <> NBAMT Then GoTo MODUPD
    ''If OEDate <> NEDate Then GoTo MODUPD
End If

'GoTo MODNOUPD

'BF_CEASEDATE was changed
MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

If xIsWorkDay = "N" Then
    rsTA("AU_BCODE") = NBCode 'clpCode(1).Text
    rsTA("AU_CEASEDATE") = rslBen("BF_CEASEDATE")
    'If OMTHCOMP <> NMTHCOMP Then rsTA("AU_MTHCCOST") = NMTHCOMP
    'If OMTHEMP <> NMTHEMP Then rsTA("AU_MTHECOST") = NMTHEMP
    'If OTAXBEN <> txtTAXBEN Then rsTA("AU_TAXBEN") = txtTAXBEN
    'If OCOVER <> NCOVER Then rsTA("AU_COVER") = NCOVER
    'If OTCOST <> NTCOST Then rsTA("AU_TCOST") = NTCOST
    'If OPremium <> lblAP Then rsTA("AU_PREMIUM") = lblAP
    'If OPPE <> NPPE Then rsTA("AU_PCE") = NPPE
    'If OPCC <> NPCC Then rsTA("AU_PCC") = NPCC
    'If OPPAMT <> NPPAMT Then
    '    rsTA("AU_PPAMT") = NPPAMT
    '    If IsNumeric(OPPAMT) Then rsTA("AU_OLDPPMT") = Val(OPPAMT)
    'End If
    'If OMAXDOL <> NMAXDOL Then rsTA("AU_MAXDOL") = NMAXDOL
    'If OEDate <> NEDate Then
    '  If IsDate(NEDate) Then
    '      rsTA("AU_EDATE") = CVDate(NEDate)
    '  End If
    'End If
    'If OPER <> txtPer Then rsTA("AU_PER") = txtPer
    'If OBAMT <> NBAMT Then rsTA("AU_BAMT") = NBAMT
    'If OUNITCOST <> medUnitCost Then rsTA("AU_UNITCOST") = IIf(medUnitCost = "", 0, medUnitCost)
    rsTA("AU_LDATE") = Date
    If IsDate(NEDate) Then 'if benefit effe date is future date, use it as LDATE
        If CVDate(NEDate) > CVDate(Date) Then
            rsTA("AU_LDATE") = CVDate(NEDate)
        End If
    End If
End If
If xIsWorkDay = "Y" Then
    rsTA("AU_LDAY") = xLastDate
    rsTA("AU_LDATE") = Date
End If
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
Else
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update
rsTA.Close

MODNOUPD:
AUDITBENF = True
Exit Function
AUDIT_ERR:

End Function

Private Sub dlpEndDate_Change()
Call WFCUpdate_Value
End Sub
