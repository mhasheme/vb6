VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPayPeriodList 
   Caption         =   "Pay Period List"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   ScaleHeight     =   4995
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCode 
      Appearance      =   0  'Flat
      DataField       =   "PP_NBR"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   480
      MaxLength       =   6
      TabIndex        =   7
      Tag             =   "01-Position's Code"
      Top             =   3300
      Width           =   975
   End
   Begin VB.TextBox txtFindKey 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      MaxLength       =   6
      TabIndex        =   6
      Tag             =   "00-Search Code"
      Top             =   3800
      Width           =   960
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
      Height          =   285
      Left            =   5730
      TabIndex        =   5
      Tag             =   "Find specific record"
      Top             =   3800
      Width           =   1200
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmPayPeriodList.frx":0000
      Height          =   3075
      Left            =   180
      OleObjectBlob   =   "frmPayPeriodList.frx":0014
      TabIndex        =   0
      Top             =   0
      Width           =   6855
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   4335
      Width           =   7335
      _Version        =   65536
      _ExtentX        =   12938
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
         Left            =   1785
         TabIndex        =   4
         Tag             =   "Print Departmental Listing"
         Top             =   135
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
         Left            =   945
         TabIndex        =   3
         Tag             =   "Close and exit this screen"
         Top             =   135
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
         Left            =   105
         TabIndex        =   2
         Tag             =   "Select this Department"
         Top             =   135
         Width           =   735
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   3660
         Top             =   180
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6450
         Top             =   150
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin INFOHR_Controls.DateLookup dlpTo 
      DataField       =   "PP_END"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   3720
      TabIndex        =   8
      Top             =   3300
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   503
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpFrom 
      DataField       =   "PP_START"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1740
      TabIndex        =   9
      Top             =   3300
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   503
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpFindFrom 
      Height          =   285
      Left            =   1740
      TabIndex        =   10
      Top             =   3800
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   503
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.DateLookup dlpFindTo 
      Height          =   285
      Left            =   3720
      TabIndex        =   11
      Top             =   3800
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      TextBoxWidth    =   1215
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frmPayPeriodList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fglbPAYP, fglbYear, fglbClosed

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdFind_Click()
Dim SQLQ As String

If Len(txtFindKey) > 0 Then
    SQLQ = "PP_NBR = " & txtFindKey
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        txtFindKey = ""
    End If
    txtFindKey.SetFocus
    Exit Sub
End If

If Len(dlpFindFrom) > 0 Then
    SQLQ = "PP_START >= " & Date_SQL(dlpFindFrom.Text)
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        dlpFindFrom = ""
    End If
    dlpFindFrom.SetFocus
    Exit Sub
End If
If Len(dlpFindTo) > 0 Then
    SQLQ = "PP_End >= " & Date_SQL(dlpFindTo.Text)
    Data1.Recordset.Requery
    Data1.Recordset.Find SQLQ
    If Data1.Recordset.EOF Then
        Data1.Refresh
    Else
        dlpFindTo = ""
    End If
    dlpFindTo.SetFocus
    Exit Sub
End If
End Sub


Private Sub cmdPrint_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = "Pay Period Listing"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub

Private Sub Form_Load()
Dim SQLQ As String

On Error GoTo Job_Err
Screen.MousePointer = HOURGLASS

Data1.ConnectionString = glbAdoIHRDB

SQLQ = "SELECT DISTINCT PP_NBR, PP_START, PP_END, PP_PAYP FROM HR_PAYPERIOD WHERE PP_YEAR=" & fglbYear
If Not IsNull(fglbClosed) Then
    SQLQ = SQLQ & " AND PP_UPLOADED = " & IIf(fglbClosed, 1, 0)
End If
SQLQ = SQLQ & " ORDER BY PP_NBR"
Data1.RecordSource = SQLQ
'Data1.RecordSource = "SELECT DISTINCT PP_NBR, PP_START, PP_END, PP_PAYP FROM HR_PAYPERIOD WHERE PP_YEAR=" & fglbYear & " ORDER BY PP_NBR"
Data1.Refresh

Screen.MousePointer = DEFAULT
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    cmdSelect.Enabled = False
End If
txtCode.Enabled = False
dlpFrom.Enabled = False
dlpTo.Enabled = False

vbxTrueGrid.Columns(3).Caption = lStr(vbxTrueGrid.Columns(3).Caption)

Exit Sub
Job_Err:
MsgBox "Error #" & Err.Number & " - " & Err.Description
Resume Next

End Sub

Private Sub Form_LostFocus()
Unload Me
End Sub

Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
    If KeyAscii = 13 Then Call cmdFind_Click
End Sub

Private Sub vbxTrueGrid_DblClick()
If Not (Data1.Recordset.EOF Or Data1.Recordset.EOF) Then
    Call cmdSelect_Click
End If
Unload Me

End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT DISTINCT PP_NBR, PP_START, PP_END, PP_PAYP FROM HR_PAYPERIOD WHERE PP_YEAR=" & fglbYear
        If Not IsNull(fglbClosed) Then
            SQLQ = SQLQ & " AND PP_UPLOADED = " & fglbClosed
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the enter key was struck
    KeyAscii = 0
    cmdClose.SetFocus
End If

End Sub

Private Sub cmdSelect_Click()
glbWeek = Data1.Recordset("PP_NBR")
glbFrom = Data1.Recordset("PP_START")
glbTo = Data1.Recordset("PP_END")
'glbTo = Data1.Recordset("PP_End")
'Ticket #29984 - Pass Pay Period code as well if Pay Period list is by Pay Period Code
If Not IsNull(Data1.Recordset("PP_PAYP")) Then
    glbPayP = Data1.Recordset("PP_PAYP")
Else
    glbPayP = ""
End If
Unload Me
End Sub

Public Property Let PayPeriodCode(vData As String)
fglbPAYP = vData
End Property

Public Property Let SelectedYear(vData As String)
fglbYear = vData
End Property

Public Property Let Closed(vData As Boolean)
fglbClosed = vData
End Property


