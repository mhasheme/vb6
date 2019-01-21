VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmVASL 
   Caption         =   "Advance Sick Leave"
   ClientHeight    =   7935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   10410
   WindowState     =   2  'Maximized
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AS_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   3090
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AS_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   2205
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "AS_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1260
      MaxLength       =   25
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtReviewDate 
      Appearance      =   0  'Flat
      DataField       =   "AS_DOA"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   285
      Left            =   1920
      MaxLength       =   10
      TabIndex        =   0
      Tag             =   "41-Date of Attendance"
      Top             =   2880
      Width           =   1215
   End
   Begin Threed.SSPanel panEEName 
      Align           =   1  'Align Top
      Height          =   435
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10410
      _Version        =   65536
      _ExtentX        =   18362
      _ExtentY        =   767
      _StockProps     =   15
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   1
      Font3D          =   1
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
         Left            =   7200
         TabIndex        =   28
         Top             =   90
         Width           =   1305
      End
      Begin VB.Label lblEENum 
         Caption         =   "Label2"
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
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Height          =   300
         Left            =   3060
         TabIndex        =   8
         Top             =   67
         Width           =   1740
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "EEId"
         DataField       =   "AS_EMPNBR"
         DataSource      =   "Data1"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   3120
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblEmpID 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         DataField       =   "AS_EMPNBR"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5760
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   5
         Top             =   105
         Width           =   1065
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   7275
      Width           =   10410
      _Version        =   65536
      _ExtentX        =   18362
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
      Begin VB.CommandButton CmdRecalc1 
         Appearance      =   0  'Flat
         Caption         =   "&Delete ASL Records for  1 Employee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   720
         TabIndex        =   13
         Tag             =   "Delete ASL Records for  1 Employee"
         Top             =   30
         Width           =   1935
      End
      Begin VB.CommandButton cmdDays 
         Appearance      =   0  'Flat
         Caption         =   "Da&ys"
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
         Left            =   3990
         TabIndex        =   12
         Tag             =   "Display Vacation and Sick Overview in Days"
         Top             =   0
         Width           =   875
      End
      Begin VB.CommandButton cmdHours 
         Appearance      =   0  'Flat
         Caption         =   "&Hours"
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
         Left            =   5010
         TabIndex        =   11
         Tag             =   "Display Vacation and Sick Overview in Hours"
         Top             =   0
         Width           =   855
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9840
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8160
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ConnectMode     =   3
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin MSMask.MaskEdBox medHours 
      DataField       =   "AS_HRSTAK"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Tag             =   "11-Hours for this reason "
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      DataField       =   "AS_HRSREP"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Tag             =   "11-Repaid Hours for this reason "
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      DataField       =   "AS_HRSOS"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Tag             =   "11-Outstanding Hours for this reason "
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medHoursDAY 
      DataField       =   "HRSTAKDAY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Tag             =   "11-Taken Hours for this reason "
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1DAY 
      DataField       =   "HRSREPDAY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Tag             =   "11-Hours for this reason "
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2DAY 
      DataField       =   "HRSOSDAY"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1920
      TabIndex        =   19
      Tag             =   "11-Hours for this reason "
      Top             =   4320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGridDAY 
      Bindings        =   "frmVASL.frx":0000
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frmVASL.frx":0014
      TabIndex        =   20
      Top             =   480
      Width           =   9375
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmVASL.frx":3AA3
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frmVASL.frx":3AB7
      TabIndex        =   21
      Top             =   480
      Width           =   9375
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   27
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Taken"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   360
      TabIndex        =   26
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "AS_COMPNO"
      DataSource      =   "Data1"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   4080
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Repaid"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   360
      TabIndex        =   24
      Top             =   3840
      Width           =   510
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Outstanding"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   360
      TabIndex        =   23
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label lblDayHrs 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DAYS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "frmVASL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xtr
Public Sub cmdCancel_Click()
    Data1.Refresh
    
End Sub

Public Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDays_Click()
  
cmdDays.Enabled = False
cmdHours.Enabled = True

medHours.Visible = False
medHoursDAY.Visible = True

MaskEdBox1.Visible = False
MaskEdBox1DAY.Visible = True

MaskEdBox2.Visible = False
MaskEdBox2DAY.Visible = True

lblDayHrs.Caption = "DAYS"

vbxTrueGridDAY.Visible = True
vbxTrueGrid.Visible = False
Call SET_UP_MODE
End Sub

Private Sub cmdHours_Click()
cmdDays.Enabled = True
cmdHours.Enabled = False

medHours.Visible = True
medHoursDAY.Visible = False

MaskEdBox1.Visible = True
MaskEdBox1DAY.Visible = False

MaskEdBox2.Visible = True
MaskEdBox2DAY.Visible = False

lblDayHrs.Caption = "HOURS"

vbxTrueGridDAY.Visible = False
vbxTrueGrid.Visible = True
Call SET_UP_MODE
End Sub




Public Sub cmdOK_Click()

    If Len(Trim(MaskEdBox1.Text)) = 0 Then
        MsgBox "Repaid cannot be blank"
        MaskEdBox1.SetFocus
        Exit Sub
    ElseIf Not IsNumeric(MaskEdBox1.Text) Then
        MsgBox "Repaid is not numeric"
        MaskEdBox1.SetFocus
        Exit Sub
    End If
    
    Dim SQLQ
    Dim dhrs
    Dim rsHREmp As New ADODB.Recordset
    Dim rsASL As New ADODB.Recordset
    
    CmdRecalc1.Enabled = False
    
    SQLQ = "SELECT ED_DHRS FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        dhrs = rsHREmp("ED_DHRS")
    Else
        dhrs = 0
    End If
    
    If lblDayHrs.Caption = "DAYS" Then
        MaskEdBox1.Text = dhrs * Val(MaskEdBox1DAY.Text)
    End If
    rsHREmp.Close
    
    SQLQ = "SELECT * FROM WHSCC_ASL "
    SQLQ = SQLQ & " WHERE AS_EMPNBR = " & glbLEE_ID & " AND AS_ATT_ID = " & Data1.Recordset("AS_ATT_ID") & " AND AS_DOA = '" & txtReviewDate & "'"
    rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsASL.EOF Then
        rsASL("AS_HRSREP") = MaskEdBox1.Text
        rsASL.Update
    End If
    rsASL.Close
    
    Data1.Refresh
    
    Call ReCalcASL(lblEEID, "")
    
    Data1.Refresh
    
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String
Me.vbxCrystal.Destination = crptToPrinter
RHeading = lblEEName & "'s Advance Sick Leave"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1
End Sub

Public Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.Destination = crptToWindow
RHeading = lblEEName & "'s Advance Sick Leave"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Private Sub CmdRecalc1_Click()
Dim SQLQ, Msg, a%
    Msg = "Are You Sure You Want To Delete All ASL Records of " & lblEEName & " ?"
    
    a% = MsgBox(Msg, 292, "Confirm Delete") '36
    If a% <> 6 Then Exit Sub
    If Len(lblEEID) = 0 Then Exit Sub
    
    CmdRecalc1.Enabled = False
    
    SQLQ = "DELETE FROM WHSCC_ASL WHERE AS_EMPNBR = " & lblEEID
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "DELETE FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND AD_REASON = 'ASL'"
    gdbAdoIhr001.Execute SQLQ
    
    SQLQ = "DELETE FROM HR_ATTENDANCE_HISTORY WHERE AH_EMPNBR = " & lblEEID
    SQLQ = SQLQ & " AND AH_REASON = 'ASL'"
    gdbAdoIhr001.Execute SQLQ
    
    Data1.Refresh
    CmdRecalc1.Enabled = True
    Call SET_UP_MODE
End Sub



Private Sub Form_Activate()
glbOnTop = "FRMVASL"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
glbOnTop = "FRMVASL"
End Sub

Private Sub Form_Load()
    
    glbOnTop = "FRMVASL"
    Data1.ConnectionString = glbAdoIHRDB

    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    
    
    If EERetrieve() = False Then
        MsgBox "Sorry, Employee can not be found"
        frmEEFIND.Show 1
    End If
    
    
    Screen.MousePointer = HOURGLASS
    
    
    If Not gSec_Upd_WHSCC_ASL Then
'        CmdRecalc1.Enabled = False

    End If
        
    Screen.MousePointer = DEFAULT

End Sub

Function EERetrieve()


Dim SQLQ As String

EERetrieve = False

Screen.MousePointer = HOURGLASS
On Error GoTo EERError

'SQLQ = "Select * from WHSCC_ASL "
'SQLQ = SQLQ & " where AS_EMPNBR = " & glbLEE_ID
'SQLQ = SQLQ & " ORDER BY AS_DOA"
SQLQ = "SELECT WHSCC_ASL.*, "
SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSTAK/ED_DHRS),2) END) AS HRSTAKDAY, "
SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSREP/ED_DHRS),2) END) AS HRSREPDAY, "
SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSOS/ED_DHRS),2) END) AS HRSOSDAY "
SQLQ = SQLQ & "FROM WHSCC_ASL INNER JOIN HREMP "
SQLQ = SQLQ & "ON WHSCC_ASL.AS_EMPNBR = HREMP.ED_EMPNBR "
SQLQ = SQLQ & " WHERE AS_EMPNBR = " & glbLEE_ID
SQLQ = SQLQ & " ORDER BY AS_DOA"

Data1.RecordSource = SQLQ
Data1.Refresh

lblEEID = glbLEE_ID
EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ASL Retrieve", "WHSCC_ASL", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


Exit Function

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
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_WHSCC_ASL
End Property

Public Property Get Addable() As Boolean

Addable = False
End Property
Public Property Get Updateble() As Boolean
If lblDayHrs.Caption = "DAYS" Then
    Updateble = False
Else
    Updateble = True
End If
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
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call modSTUPD(TF)
End Sub
Private Sub modSTUPD(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If


fUPMode = TF    ' update mode

'If Not lblDayHrs.Caption = "DAYS" Then
'    MaskEdBox1.Enabled = True
'    MaskEdBox1.SetFocus
'End If

MaskEdBox1.Enabled = TF
CmdRecalc1.Enabled = TF

If lblDayHrs.Caption = "DAYS" Then
    cmdDays.Enabled = False
    cmdHours.Enabled = True
Else
    cmdHours.Enabled = False
    cmdDays.Enabled = True
End If

End Sub
Private Sub lblEEID_Change()
lblEENum = lblEEID
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "ASL - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End If
End Sub



Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT WHSCC_ASL.*, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSTAK/ED_DHRS),2) END) AS HRSTAKDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSREP/ED_DHRS),2) END) AS HRSREPDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSOS/ED_DHRS),2) END) AS HRSOSDAY "
        SQLQ = SQLQ & "FROM WHSCC_ASL INNER JOIN HREMP "
        SQLQ = SQLQ & "ON WHSCC_ASL.AS_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE AS_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub



Private Sub vbxTrueGridDAY_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGridDAY.Tag = "ASC" Then
            vbxTrueGridDAY.Tag = "DESC"
        Else
            vbxTrueGridDAY.Tag = "ASC"
        End If
        
        SQLQ = "SELECT WHSCC_ASL.*, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSTAK/ED_DHRS),2) END) AS HRSTAKDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSREP/ED_DHRS),2) END) AS HRSREPDAY, "
        SQLQ = SQLQ & "(CASE WHEN ED_DHRS=0 THEN 0 ELSE ROUND((AS_HRSOS/ED_DHRS),2) END) AS HRSOSDAY "
        SQLQ = SQLQ & "FROM WHSCC_ASL INNER JOIN HREMP "
        SQLQ = SQLQ & "ON WHSCC_ASL.AS_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE AS_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGridDAY.Columns(ColIndex).DataField & " " & vbxTrueGridDAY.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
End Sub
