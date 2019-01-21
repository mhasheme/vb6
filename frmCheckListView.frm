VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckListView 
   Caption         =   "Check ListView"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15510
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   9405
   ScaleWidth      =   15510
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraReptEmpLis2 
      Height          =   3735
      Left            =   4080
      TabIndex        =   37
      Top             =   5040
      Visible         =   0   'False
      Width           =   11175
      Begin INFOHR_Controls.EmployeeLookup elpRept 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   39
         Tag             =   "10-Reporting Authority"
         Top             =   3000
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.Frame frmAT 
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   2040
         TabIndex        =   44
         Top             =   3240
         Visible         =   0   'False
         Width           =   5115
         Begin VB.OptionButton optAT 
            Caption         =   "Active Employee"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   46
            Top             =   150
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton optAT 
            Caption         =   "Terminated Employee"
            Height          =   255
            Index           =   1
            Left            =   2490
            TabIndex        =   45
            Top             =   150
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdOne 
         Appearance      =   0  'Flat
         Caption         =   "Save Change"
         Height          =   375
         Left            =   8040
         TabIndex        =   43
         Tag             =   "Save the changes made"
         Top             =   2920
         Width           =   1275
      End
      Begin VB.CommandButton cmdAll 
         Appearance      =   0  'Flat
         Caption         =   "All Employees"
         Height          =   375
         Left            =   9360
         TabIndex        =   42
         Tag             =   "Save the changes made"
         Top             =   2920
         Width           =   1275
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "frmCheckListView.frx":0000
         Height          =   2625
         Left            =   120
         OleObjectBlob   =   "frmCheckListView.frx":0014
         TabIndex        =   38
         Top             =   240
         Width           =   10755
      End
      Begin VB.Label lblRAName 
         Caption         =   "Label1"
         Height          =   255
         Left            =   7920
         TabIndex        =   41
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lblEmpNum2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "New Reporting Authority #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   3000
         Width           =   1740
      End
   End
   Begin VB.Frame fraReptEmpList 
      DragMode        =   1  'Automatic
      Height          =   735
      Left            =   6600
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   7695
      Begin INFOHR_Controls.EmployeeLookup elpRept 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   33
         Tag             =   "10-Reporting Authority"
         Top             =   240
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.Label lblStDate 
         Caption         =   "Label1"
         Height          =   255
         Left            =   6600
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label lblEmpNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Reporting Authority #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   34
         Top             =   285
         Width           =   1860
      End
   End
   Begin VB.Frame fraImpCurrencyExch 
      Height          =   3615
      Left            =   6600
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   3720
         TabIndex        =   23
         Tag             =   "Save the changes made"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdImp 
         Appearance      =   0  'Flat
         Caption         =   "&Import"
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Tag             =   "Save the changes made"
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox txtUS 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   19
         Tag             =   "00-US Cell"
         Text            =   "B"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtCDN 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   18
         Tag             =   "00-CDN Cell"
         Text            =   "B"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtMTH 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   6240
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "MTH"
         Top             =   240
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.ComboBox ComMTH 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3480
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Tag             =   "01-Month"
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   6480
         TabIndex        =   21
         Tag             =   "Click to select the location"
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
         Height          =   315
         Left            =   1440
         TabIndex        =   20
         Tag             =   "00-File Name (Include Extension TXT)"
         Top             =   1680
         Width           =   4905
      End
      Begin MSComDlg.CommonDialog AttachmentDialog 
         Left            =   7320
         Top             =   1560
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Filter          =   "*.xls;*.xlsx"
      End
      Begin MSMask.MaskEdBox MskFiscalYear 
         DataField       =   "IP_YEAR"
         Height          =   315
         Left            =   1440
         TabIndex        =   16
         Tag             =   "01-High Dollars"
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "###0"
         PromptChar      =   "_"
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Codes must start from cell B24"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   2760
         TabIndex        =   30
         Top             =   1200
         Width           =   2805
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Currency Codes must start from cell B2"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   29
         Top             =   720
         Width           =   3555
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "US Cell"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   28
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "CDN Cell"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   27
         Top             =   765
         Width           =   975
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   2760
         TabIndex        =   26
         Top             =   260
         Width           =   450
      End
      Begin VB.Label lblCriteria 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   260
         Width           =   330
      End
      Begin VB.Label lblCriteria 
         BackStyle       =   0  'Transparent
         Caption         =   "Import From"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   1740
         Width           =   1260
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Select"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Question"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   2
      EndProperty
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   8745
      Width           =   15510
      _Version        =   65536
      _ExtentX        =   27358
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
      Begin VB.CommandButton cmdCancelTran 
         Appearance      =   0  'Flat
         Caption         =   "Cancel Transaction"
         Height          =   375
         Left            =   8520
         TabIndex        =   11
         Tag             =   "Save the changes made"
         Top             =   120
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "Print"
         Height          =   375
         Left            =   7440
         TabIndex        =   10
         Tag             =   "Save the changes made"
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdView 
         Appearance      =   0  'Flat
         Caption         =   "View"
         Height          =   375
         Left            =   6360
         TabIndex        =   9
         Tag             =   "Save the changes made"
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdClear 
         Appearance      =   0  'Flat
         Caption         =   "Clear All"
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Tag             =   "Save the changes made"
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdSele 
         Appearance      =   0  'Flat
         Caption         =   "Select All"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Tag             =   "Save the changes made"
         Top             =   120
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   2520
         TabIndex        =   6
         Tag             =   "Cancel the changes made"
         Top             =   120
         Width           =   915
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   11880
         Top             =   0
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   14520
         Top             =   0
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
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataSource      =   "data1"
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Tag             =   "00-Region"
      Top             =   3960
      Visible         =   0   'False
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin MSMask.MaskEdBox MskTPerc 
      DataSource      =   "data1"
      Height          =   315
      Left            =   2475
      TabIndex        =   3
      Tag             =   "01-MidPoint Dollars"
      Top             =   4440
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskFiscalYea2 
      DataField       =   "IP_YEAR"
      Height          =   315
      Left            =   2475
      TabIndex        =   1
      Tag             =   "01-High Dollars"
      Top             =   4920
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "###0"
      PromptChar      =   "_"
   End
   Begin VB.Label lblNote1 
      Caption         =   "The following employees have <<Employee>> as their Reporting Authority #1. Who should be their Interim/New Reporting Authority?"
      Height          =   615
      Left            =   6600
      TabIndex        =   35
      Top             =   4440
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   600
      TabIndex        =   31
      Top             =   4920
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   600
      TabIndex        =   13
      Top             =   4440
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lbltitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmCheckListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xSeleList(200)
Dim xSeleTot As Integer
Dim ImportFile
Dim xCDNStartLine
Dim xUSStartLine
Dim xOldReptNo, xNewReptNo
Dim xLocRptName

Private Sub cmdAll_Click()
    If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
        If Len(elpRept(1).Text) = 0 Then
            MsgBox "Reporting Authority # is required"
            elpRept(1).SetFocus
            Exit Sub
        End If
        If Not elpRept(1).ListChecker Then
            'MsgBox "Invalid Reporting Authority #"
            elpRept(1).SetFocus
            Exit Sub
        End If
        
        If optAT(0).Value Then
            If IsNumeric(elpRept(1).Text) Then
                If IsRept1PosNotMatchEmpRept1(Data1.Recordset("BM_WAITPERIOD"), elpRept(1).Text) Then
                    'continue
                Else
                    Exit Sub
                End If
            End If
        End If
        
        xID = Data1.Recordset("BM_BENE_ID")
        SQLQ = "UPDATE HRBENGRPLIST SET BM_WRKID = " & elpRept(1).Text & " WHERE BM_WRKEMP = '" & glbUserID & "' " 'AND BM_BENE_ID = " & xID & " "
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "UPDATE HRBENGRPLIST SET BM_COMMENTS = '" & Left(elpRept(1).Caption, 50) & "' WHERE BM_WRKEMP = '" & glbUserID & "' " ' AND BM_BENE_ID = " & xID & " "
        gdbAdoIhr001.Execute SQLQ
        Data1.Refresh
        Data1.Recordset.Find "BM_BENE_ID = " & xID & " "
    End If
End Sub

Private Sub cmdBrowse_Click()
AttachmentDialog.DialogTitle = "Select the file to import..."
AttachmentDialog.Filter = "*.xls;*.xlsx|*.xls;*.xlsx"
AttachmentDialog.FilterIndex = 1
AttachmentDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
AttachmentDialog.ShowOpen
If Len(AttachmentDialog.FileName) <> 0 Then
    txtFileName.Text = AttachmentDialog.FileName
End If
End Sub

Private Sub cmdCancel_Click()
'lvw.ListItems.Clear 'how to clean all items
    glbWFC_IncePlanID = -1
    Unload Me
End Sub

Private Sub cmdCancelTran_Click()
    glbWFC_IncePlanID = -1
    glbWFC_CancelTransaction = True
    Unload Me
End Sub

Private Sub cmdClear_Click()
Dim xTot As Integer
Dim I As Integer
Dim xList As String

    xTot = lvw.ListItems.count
    
    xSeleTot = 1
    xList = ""
    For I = 1 To xTot
        lvw.ListItems(I).Checked = False
    Next
End Sub

Private Sub cmdClose_Click()
    glbWFC_IncePlanID = -1
    Unload Me
End Sub

Private Sub cmdImp_Click()
Dim Msg As String, a%
    If Not chkIMPCurrency() Then Exit Sub
    
    Msg = "Are you sure you want to Import Currency Exchange from "
    Msg = Msg & Chr(10) & txtFileName.Text & "?"
    
    a% = MsgBox(Msg, 36, "Confirm Import")
    If a% <> 6 Then
        Exit Sub
    End If

    Call WFCIPImpCurrency
    
    MsgBox "   Finished.   "
    Unload Me
    
End Sub

Private Function chkIMPCurrency()
Dim X%, Y%
Dim xStr

chkIMPCurrency = False

On Error GoTo chkIMPCurrency_Err

If Len(MskFiscalYear.Text) > 0 Then
    If Not IsNumeric(MskFiscalYear.Text) Then
        MsgBox "Invalid Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
    If Not Len(MskFiscalYear.Text) = 4 Then
        MsgBox "Invalid Year."
        MskFiscalYear.SetFocus
        Exit Function
    End If
Else
    MsgBox "Fiscal Year is a required field"
    MskFiscalYear.SetFocus
    Exit Function
End If
If Len(ComMTH.Text) = 0 Then
    MsgBox "Month is a required field"
    ComMTH.SetFocus
    Exit Function
End If


If Len(txtCDN.Text) = 0 Then
    MsgBox "CDN Cell is a required field"
    txtCDN.SetFocus
    Exit Function
End If
If Not UCase(Left(txtCDN.Text, 1)) = "B" Then
    MsgBox "The first character of CDN Cell must B"
    txtCDN.SetFocus
    Exit Function
End If
xStr = Mid(txtCDN.Text, 2, 5)
If Not IsNumeric(xStr) Then
    MsgBox "Invalid CDN Cell, you must enter a number after letter B"
    txtCDN.SetFocus
    Exit Function
End If
xCDNStartLine = xStr

If Len(txtUS.Text) = 0 Then
    MsgBox "US Cell is a required field"
    txtUS.SetFocus
    Exit Function
End If
If Not UCase(Left(txtUS.Text, 1)) = "B" Then
    MsgBox "The first character of US Cell must B"
    txtUS.SetFocus
    Exit Function
End If
xStr = Mid(txtUS.Text, 2, 5)
If Not IsNumeric(xStr) Then
    MsgBox "Invalid US Cell, you must enter a number after letter B"
    txtUS.SetFocus
    Exit Function
End If
xUSStartLine = xStr

If Len(txtFileName.Text) = 0 Then
    MsgBox "Please enter Import From File"
    txtFileName.SetFocus
    Exit Function
End If

ImportFile = txtFileName.Text

If Dir(ImportFile) = "" Then
  MsgBox ImportFile & " File not Found."
  txtFileName.SetFocus
  Exit Function
End If
    

chkIMPCurrency = True

Exit Function

chkIMPCurrency_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkIMPCurrency", "HRIP_CURRENCY_EXCHG", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Private Sub cmdOK_Click()
Dim Msg As String, a%
Dim xTot As Integer
Dim I As Integer
Dim xList As String
Dim xDisPer
Dim xStartDate

If glbWFC_IPPopFormName = "WIADivisoinList" Then
    xTot = lvw.ListItems.count
    
    xSeleTot = 1
    xList = ""
    For I = 1 To xTot
        If lvw.ListItems(I).Checked Then
            xList = xList & "'" & lvw.ListItems(I).SubItems(2) & "',"
            xSeleList(xSeleTot) = lvw.ListItems(I).SubItems(2)
            xSeleTot = xSeleTot + 1
        End If
    Next
    
    'MsgBox xList
    If Len(xList) = 0 Then
        MsgBox "No Division Selected"
        Exit Sub
    End If
    xSeleTot = xSeleTot - 1

    Msg = "Are you sure you want to copy this record to all the Plants which you selected?"
    'Msg = Msg & Chr(10) & "This Province?"
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If

    Call WFCIPFactorRepeatFunc
    
    Unload Me
End If

If glbWFC_IPPopFormName = "WFCEmpListWithOldPos" Then 'Ticket #29183 Franks 09/13/2016
    xTot = lvw.ListItems.count
    If xTot = 0 Then
        MsgBox "There is no employee whose position is '" & glbChgTermReason & "' "
        Unload Me
    End If
    
    xSeleTot = 1
    xList = ""
    For I = 1 To xTot
        If lvw.ListItems(I).Checked Then
            xList = xList & "'" & lvw.ListItems(I).SubItems(1) & "',"
            xSeleList(xSeleTot) = lvw.ListItems(I).SubItems(1)
            xSeleTot = xSeleTot + 1
        End If
    Next
    
    'MsgBox xList
    If Len(xList) = 0 Then
        MsgBox "No Employee Selected"
        Exit Sub
    End If
    xSeleTot = xSeleTot - 1
    
    Msg = "This function will do the following tasks for the selected employees:" & Chr(10) & Chr(10)
    Msg = Msg & "1. Do a 'New Position/Same Salary' on Employee Position screen. Position Start Date will equal the Start Date of the new Position Master record " & Chr(10)
    Msg = Msg & "2. A salary record is created with the new Position and Position Start Date. " & Chr(10)
    Msg = Msg & "3. Update the Audit Master. " & Chr(10) & Chr(10)
    'Msg = Msg & "4. Update Salary Dependent Benefits. " & Chr(10) & Chr(10)
    
    Msg = Msg & "Are you sure you want to do these?"
    'Msg = Msg & Chr(10) & "This Province?"
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        'Exit Sub
        glbWFC_IncePlanID = -1
        Unload Me
    End If
    
    'xStartDate = GetJobData(glbChgTermReason, "JB_SDATE", Date) '??? glbChgTermReason is not new position 'glbSPCPPay
    'Call WFCEmpNewPosUptFunc(glbChgTermReason, xStartDate)
    xStartDate = GetJobData(glbSPCPPay, "JB_SDATE", Date)
    Call WFCEmpNewPosUptFunc(glbSPCPPay, xStartDate)

    Unload Me

End If

If glbWFC_IPPopFormName = "UpdateBUFin" Then
    If Len(clpCode(0).Text) = 0 Then
        MsgBox lbltitle(0).Caption & " is a required field"
        clpCode(0).SetFocus
        Exit Sub
    End If
    If Len(clpCode(0).Text) > 0 And clpCode(0).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(0).SetFocus
        Exit Sub
    End If
    If Len(MskTPerc.Text) = 0 Then
        MsgBox "Percentage is a required field"
        MskTPerc.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(MskTPerc.Text) Then
        MsgBox "Invalid Percentage"
        MskTPerc.SetFocus
        Exit Sub
    End If
    
    xDisPer = Val(MskTPerc.Text) * 100
    xDisPer = xDisPer & "%"
    
    Msg = "Are you sure you want to Update BU Fin with " & xDisPer & " for " & clpCode(0).Text & " and " & xExpYears & "?"
    'Msg = Msg & Chr(10) & "This Province?"
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If
        
    'Call WFCIPUpdateBUFinFunc
    SQLQ = "UPDATE HRIP_FACTORS SET IP_A_BU_FIN = " & MskTPerc.Text & " WHERE IP_YEAR = " & xExpYears & " "
    SQLQ = SQLQ & "AND IP_REGION = '" & clpCode(0).Text & "' "
    gdbAdoIhr001.Execute SQLQ, I
    
    If I = 0 Then
        MsgBox "No record updated. "
        glbWFC_IncePlanID = -1
    Else
        MsgBox I & " records updated. "
    End If
    
    Unload Me
End If

If glbWFC_IPPopFormName = "UpdateCorpFin" Then
    If Len(MskTPerc.Text) = 0 Then
        MsgBox "Percentage is a required field"
        MskTPerc.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(MskTPerc.Text) Then
        MsgBox "Invalid Percentage"
        MskTPerc.SetFocus
        Exit Sub
    End If
    
    xDisPer = Val(MskTPerc.Text) * 100
    xDisPer = xDisPer & "%"
    
    Msg = "Are you sure you want to Update Corp Fin with " & xDisPer & " for " & xExpYears & "?"
    'Msg = Msg & Chr(10) & "This Province?"
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If
        
    'Call WFCIPUpdateBUFinFunc
    SQLQ = "UPDATE HRIP_FACTORS SET IP_A_CORP_FIN = " & MskTPerc.Text & " WHERE IP_YEAR = " & xExpYears & " "
    gdbAdoIhr001.Execute SQLQ, I
    
    If I = 0 Then
        MsgBox "No record updated. "
        glbWFC_IncePlanID = -1
    Else
        MsgBox I & " records updated. "
    End If
    
    Unload Me
End If

If glbWFC_IPPopFormName = "UpdateSalesComm" Then
    If Len(MskTPerc.Text) = 0 Then
        MsgBox "Percentage is a required field"
        MskTPerc.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(MskTPerc.Text) Then
        MsgBox "Invalid Percentage"
        MskTPerc.SetFocus
        Exit Sub
    End If
    
    xDisPer = Val(MskTPerc.Text) * 100
    xDisPer = xDisPer & "%"
    
    Msg = "Are you sure you want to Update Sales Comm with " & xDisPer & " for " & xExpYears & "?"
    'Msg = Msg & Chr(10) & "This Province?"
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If

    SQLQ = "UPDATE HRIP_FACTORS SET IP_A_SALES_COMM = " & MskTPerc.Text & " WHERE IP_YEAR = " & xExpYears & " "
    gdbAdoIhr001.Execute SQLQ, I
    
    If I = 0 Then
        MsgBox "No record updated. "
        glbWFC_IncePlanID = -1
    Else
        MsgBox I & " records updated. "
    End If
    
    Unload Me
End If

If glbWFC_IPPopFormName = "UpdateROIC" Then
    
    If Len(MskFiscalYea2.Text) > 0 Then
        If Not IsNumeric(MskFiscalYea2.Text) Then
            MsgBox "Invalid Year."
            MskFiscalYea2.SetFocus
            Exit Sub
        End If
        If Not Len(MskFiscalYea2.Text) = 4 Then
            MsgBox "Invalid Year."
            MskFiscalYea2.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Fiscal Year is a required field"
        MskFiscalYea2.SetFocus
        Exit Sub
    End If

    If Len(MskTPerc.Text) = 0 Then
        MsgBox lbltitle(1).Caption & " is a required field"
        MskTPerc.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(MskTPerc.Text) Then
        MsgBox "Invalid " & lbltitle(1).Caption
        MskTPerc.SetFocus
        Exit Sub
    End If
    
    xDisPer = Val(MskTPerc.Text)
    'xDisPer = xDisPer & "%"
    
    'Msg = "Are you sure you want to Update ROIC with " & xDisPer & " for " & xExpYears & "?"
    Msg = "Are you sure you want to Update ROIC with " & xDisPer & " for " & MskFiscalYea2.Text & "?"
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
    End If

    SQLQ = "UPDATE HRIP_FACTORS SET IP_ROIC = " & MskTPerc.Text & " WHERE IP_YEAR = " & MskFiscalYea2.Text & " "
    gdbAdoIhr001.Execute SQLQ, I
    
    If I = 0 Then
        MsgBox "No record updated. "
        glbWFC_IncePlanID = -1
    Else
        MsgBox I & " records updated. "
    End If
    
    Unload Me
End If

'Ticket #29438 Franks 11/08/2016
If glbWFC_IPPopFormName = "WFCEmpListWithRept" Then
    
    'check if there is change on RA #
    If Not IsRAChanged Then
        MsgBox "Can not do Update since there is no change on Reporting Authority #"
        Exit Sub
    End If
    
    'Msg = "The Reporting Authority # has been changed " & Chr(10) & Chr(10) '" from " & xOldReptNo & " to " & elpRept(0).Text & Chr(10) & Chr(10)
    Msg = "Have you made New RA# changes and clicked on either Save Change or All Employees? " & Chr(10) & Chr(10)
    Msg = Msg & "Click Yes to finish this change. " & Chr(10) & Chr(10)
    Msg = Msg & "Click No without change. " '& Chr(10)
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
        'glbWFC_IncePlanID = -1
        'Unload Me
    End If
    If a% = 6 Then 'Update
        Call WFCEmpReptUptFunc
        glbWFC_IncePlanID = -1
        Unload Me
    End If

End If
    
'Ticket #29438 Franks 11/08/2016
If glbWFC_IPPopFormName = "WFCEmpListWithRepTerm" Or glbWFC_IPPopFormName = "WFCEmpListWithRepTran" Then
    'check if there is change on RA #
    
    'If Not IsRAChanged Then
    '    MsgBox "Can not do Update since there is no change on Reporting Authority #"
    '    Exit Sub
    'End If
    
    If IsNewRABlank Then
        MsgBox "Can not leave New RA # as blank"
        Exit Sub
    End If
    
    If IsNewRASameAsOldRA Then
        MsgBox "New RA # can not be same as Current RA#"
        Exit Sub
    End If
    

    'Msg = "The Reporting Authority # has been changed " & Chr(10) & Chr(10) '" from " & xOldReptNo & " to " & elpRept(0).Text & Chr(10) & Chr(10)
    'Msg = Msg & "Click Yes to finish this change. " & Chr(10) & Chr(10)
    'Msg = Msg & "Click No without change. " '& Chr(10)
    

    Msg = "Have you made all New RA# changes and clicked on either Save Change or All Employees? " & Chr(10) & Chr(10) '" from " & xOldReptNo & " to " & elpRept(0).Text & Chr(10) & Chr(10)
    Msg = Msg & "Click Yes to finish this change. " & Chr(10) & Chr(10)
    Msg = Msg & "Click No without change. " '& Chr(10)
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
        'glbWFC_IncePlanID = -1
        'Unload Me
    End If
    If a% = 6 Then 'Update
        Call WFCEmpReptUptFunc
        glbWFC_IncePlanID = -1
        Unload Me
    End If
End If

If glbWFC_IPPopFormName = "WFCEmpListWithRepTranIn" Then
    'check if there is change on RA #
    
    If IsNewRABlank Then
        MsgBox "Can not leave New RA # as blank"
        Exit Sub
    End If
    
    If Not IsRAChanged Then
        MsgBox "Can not do Update since there is no change on Reporting Authority #"
        Exit Sub
    End If
    
    'If IsNewRASameAsOldRA Then
    '    MsgBox "New RA # can not be same as Current RA#"
    '    Exit Sub
    'End If
   
    Msg = "Have you made New RA# changes and clicked on either Save Change or All Employees? " & Chr(10) & Chr(10) '" from " & xOldReptNo & " to " & elpRept(0).Text & Chr(10) & Chr(10)
    Msg = Msg & "Click Yes to finish this change. " & Chr(10) & Chr(10)
    Msg = Msg & "Click No without change. " '& Chr(10)
    
    a% = MsgBox(Msg, 36, "Confirm")
    If a% <> 6 Then
        Exit Sub
        'glbWFC_IncePlanID = -1
        'Unload Me
    End If
    If a% = 6 Then 'Update
        Call WFCEmpReptUptFunc
        glbWFC_IncePlanID = -1
        Unload Me
    End If
End If

End Sub

Private Sub WFCIPImpCurrency()
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsIPFactors As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim xDiv, xPlant
Dim SQLQ As String
Dim I As Integer
Dim xCode, xRate
Dim xYear, xMTHSEQ, xMTHDesc, xCONVERT_NO, xORDER, xCURRENCYIND1, xCURRENCYIND2, xCOUNTRY, xCURRENCYINDF

    Screen.MousePointer = vbHourglass
    
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = "Please Wait..."
    MDIMain.panHelp(2).Caption = ""
    
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(ImportFile)
    Set exSheet = exBook.Worksheets(1)
        
    xCURRENCYIND1 = "CAD"
    xCURRENCYIND2 = "USD"
    'CDN Import to Convert 1 - Start
    For I = 0 To 19
        xCode = exSheet.Cells(2, I + 2)
        If Len(xCode) > 0 Then xCode = Left(Trim(xCode), 4)
        'get rate
        xRate = 0
        If Len(xCode) > 0 Then
            Call CheckHRTABLCode("WFCI", xCode, xCode & " - IMPORT")
            xRate = exSheet.Cells(xCDNStartLine, I + 2)
            If Not IsNumeric(xRate) Then xRate = 0
            'add to Currency Table
            xYear = MskFiscalYear.Text
            xMTHSEQ = Left(ComMTH.Text, 2)
            xMTHDesc = Left(Mid(ComMTH.Text, 4, 30), 30)
            xCONVERT_NO = 1
            xORDER = I
            xCOUNTRY = ""
            xCURRENCYINDF = xCode
            Call WFCImpCurrencyFunc(xYear, xMTHSEQ, xMTHDesc, xCONVERT_NO, xORDER, xCURRENCYIND1, xCURRENCYIND2, xCOUNTRY, xCURRENCYINDF, xRate)
        End If
    Next
    'CDN Import to Convert 1 - end
    
    'USD Import to Convert 2 - Start
    For I = 0 To 19
        xCode = exSheet.Cells(24, I + 2)
        If Len(xCode) > 0 Then xCode = Left(Trim(xCode), 4)
        'get rate
        xRate = 0
        If Len(xCode) > 0 Then
            Call CheckHRTABLCode("WFCI", xCode, xCode & " - IMPORT")
            xRate = exSheet.Cells(xUSStartLine, I + 2)
            If Not IsNumeric(xRate) Then xRate = 0
            'add to Currency Table
            xYear = MskFiscalYear.Text
            xMTHSEQ = Left(ComMTH.Text, 2)
            xMTHDesc = Left(Mid(ComMTH.Text, 4, 30), 30)
            xCONVERT_NO = 2
            xORDER = I
            xCOUNTRY = ""
            xCURRENCYINDF = xCode
            Call WFCImpCurrencyFunc(xYear, xMTHSEQ, xMTHDesc, xCONVERT_NO, xORDER, xCURRENCYIND1, xCURRENCYIND2, xCOUNTRY, xCURRENCYINDF, xRate)
        End If
    Next
    'USD Import to Convert 1 - end
    
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    Screen.MousePointer = vbDefault
        

    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""

End Sub

Private Sub WFCEmpNewPosUptFunc(xJob, xStartDate)
Dim rsEmpPos As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset
Dim xEmpNo, xPlant
Dim SQLQ As String
Dim I As Integer

Screen.MousePointer = HOURGLASS


    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001W.CommitTrans
                
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    'SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_ID = " & glbWFC_IncePlanID & " "
    'rsIPFactors.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'If Not rsIPFactors.EOF Then
    For I = 1 To xSeleTot
        xEmpNo = xSeleList(I)
        
        SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 "
        SQLQ = SQLQ & "AND JH_REPTAU = " & xEmpNo & " "
        If rsTmp.State <> 0 Then rsTmp.Close
        rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTmp.EOF Then 'this employee is RA1 then add to this list
            'Add to the temp table
            rsEListWRK.AddNew
            rsEListWRK("TT_COMPNO") = "001"
            rsEListWRK("TT_EMPNBR") = xEmpNo
            rsEListWRK("TT_WRKEMP") = glbUserID
            rsEListWRK.Update
        End If
        
        'find the current one
        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 AND JH_EMPNBR = " & xEmpNo & " "
        If rsEmpPos.State <> 0 Then rsEmpPos.Close
        rsEmpPos.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEmpPos.EOF Then
            GoTo next_emp 'skip if no current position found
        End If
        
        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 AND JH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND JH_JOB = '" & xJob & "' "
        SQLQ = SQLQ & "AND JH_SDATE = " & Date_SQL(xStartDate) & " "
        
        If rsAdd.State <> 0 Then rsAdd.Close
        rsAdd.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsAdd.EOF Then
            GoTo next_emp 'don't update it if found duplicated record
        End If
        rsAdd.AddNew
        rsAdd("JH_EMPNBR") = rsEmpPos("JH_EMPNBR")
        rsAdd("JH_SDATE") = xStartDate
        rsAdd("JH_CURRENT") = 1
        rsAdd("JH_JOB") = xJob
        rsAdd("JH_JREASON") = "TITL" 'rsEmpPos("JH_JREASON")
        rsAdd("JH_REPTAU") = rsEmpPos("JH_REPTAU")
        rsAdd("JH_DHRS") = rsEmpPos("JH_DHRS")
        rsAdd("JH_WHRS") = rsEmpPos("JH_WHRS")
        rsAdd("JH_PHRS") = rsEmpPos("JH_PHRS")
        rsAdd("JH_SHIFT") = rsEmpPos("JH_SHIFT")
        rsAdd("JH_FTENUM") = rsEmpPos("JH_FTENUM")
        rsAdd("JH_FTEHRS") = rsEmpPos("JH_FTEHRS")
        rsAdd("JH_ORG") = rsEmpPos("JH_ORG")
        rsAdd("JH_PT") = rsEmpPos("JH_PT")
        rsAdd("JH_DIV") = rsEmpPos("JH_DIV")
        rsAdd("JH_DEPTNO") = rsEmpPos("JH_DEPTNO")
        rsAdd("JH_SECTION") = rsEmpPos("JH_SECTION")
        rsAdd("JH_PAYROLL_ID") = rsEmpPos("JH_PAYROLL_ID")
        rsAdd("JH_GRID") = rsEmpPos("JH_GRID")
        rsAdd("JH_REPTAU4") = rsEmpPos("JH_REPTAU4")
        rsAdd("JH_LDATE") = Date
        rsAdd("JH_LTIME") = Time$
        rsAdd("JH_LUSER") = glbUserID
        rsAdd.Update
        
        Call WFC_AUDITEMPPOS(xEmpNo, xJob, xStartDate)
        
        'resetup the previous current record with end date
        'call Set_Current_Flag
        rsEmpPos("JH_CURRENT") = 0
        If IsDate(glbChgTermDate) Then
            rsEmpPos("JH_ENDDATE") = CVDate(glbChgTermDate)
        End If
        rsEmpPos.Update
        
        'update salary
        Call Upd_Related_Salary_public(xEmpNo, rsAdd, "N")
        
next_emp:

    Next
    'End If
    
    rsEListWRK.Close
    
Screen.MousePointer = DEFAULT

End Sub

Private Sub WFCIPFactorRepeatFunc()
Dim rsIPFactors As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim xDiv, xPlant
Dim SQLQ As String
Dim I As Integer

    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_ID = " & glbWFC_IncePlanID & " "
    rsIPFactors.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsIPFactors.EOF Then
        For I = 1 To xSeleTot
            xDiv = xSeleList(I)
            xPlant = ""
            SQLQ = "SELECT DIV,DV_SECTION FROM HR_DIVISION WHERE DIV = '" & xDiv & "' "
            If rsAdd.State <> 0 Then rsAdd.Close
            rsAdd.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsAdd.EOF Then
                If Not IsNull(rsAdd("DV_SECTION")) Then
                    xPlant = rsAdd("DV_SECTION")
                End If
            End If
            rsAdd.Close
            
            If Not IsNull(rsIPFactors("IP_SECTION")) Then
                If Len(xPlant) > 0 Then
                    SQLQ = "SELECT * FROM HRIP_FACTORS WHERE IP_YEAR = " & rsIPFactors("IP_YEAR") & " "
                    SQLQ = SQLQ & "AND IP_POSTYPE = '" & rsIPFactors("IP_POSTYPE") & "' "
                    SQLQ = SQLQ & "AND IP_SECTION = '" & xPlant & "' "
                    If rsAdd.State <> 0 Then rsAdd.Close
                    rsAdd.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsAdd.EOF Then 'not found the matching record for this plant, then add it
                        rsAdd.AddNew
                        rsAdd("IP_YEAR") = rsIPFactors("IP_YEAR")
                        rsAdd("IP_POSTYPE") = rsIPFactors("IP_POSTYPE")
                        rsAdd("IP_REGION") = rsIPFactors("IP_REGION")
                        rsAdd("IP_SECTION") = xPlant
                        rsAdd("IP_T_PLANT_OBJ") = rsIPFactors("IP_T_PLANT_OBJ")
                        rsAdd("IP_T_BU_FIN") = rsIPFactors("IP_T_BU_FIN")
                        rsAdd("IP_T_CORP_FIN") = rsIPFactors("IP_T_CORP_FIN")
                        rsAdd("IP_T_SALES_IND") = rsIPFactors("IP_T_SALES_IND")
                        rsAdd("IP_T_SALES_COMM") = rsIPFactors("IP_T_SALES_COMM")
                        rsAdd("IP_T_CORP_OBJ") = rsIPFactors("IP_T_CORP_OBJ")
                        rsAdd("IP_A_PLANT_OBJ") = rsIPFactors("IP_A_PLANT_OBJ")
                        rsAdd("IP_A_BU_FIN") = rsIPFactors("IP_A_BU_FIN")
                        rsAdd("IP_A_CORP_FIN") = rsIPFactors("IP_A_CORP_FIN")
                        rsAdd("IP_A_SALES_IND") = rsIPFactors("IP_A_SALES_IND")
                        rsAdd("IP_A_SALES_COMM") = rsIPFactors("IP_A_SALES_COMM")
                        rsAdd("IP_A_CORP_OBJ") = rsIPFactors("IP_A_CORP_OBJ")
                        rsAdd("IP_ROIC") = rsIPFactors("IP_ROIC")
                        rsAdd("IP_LDATE") = Date
                        rsAdd("IP_LTIME") = Time$
                        rsAdd("IP_LUSER") = glbUserID
                        rsAdd.Update
                    End If
                End If
            End If
        Next
    End If

End Sub

Private Sub cmdOne_Click()
Dim SQLQ
Dim xID
    
    If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
        If Len(elpRept(1).Text) = 0 Then
            MsgBox "Reporting Authority # is required"
            elpRept(1).SetFocus
            Exit Sub
        End If
        If Not elpRept(1).ListChecker Then
            'MsgBox "Invalid Reporting Authority #"
            elpRept(1).SetFocus
            Exit Sub
        End If
        
        If optAT(0).Value Then
            If IsNumeric(elpRept(1).Text) Then
                If IsRept1PosNotMatchEmpRept1(Data1.Recordset("BM_WAITPERIOD"), elpRept(1).Text) Then
                    'continue
                Else
                    Exit Sub
                End If
            End If
        End If
        
        xID = Data1.Recordset("BM_BENE_ID")
        SQLQ = "UPDATE HRBENGRPLIST SET BM_WRKID = " & elpRept(1).Text & " WHERE BM_WRKEMP = '" & glbUserID & "' AND BM_BENE_ID = " & xID & " "
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "UPDATE HRBENGRPLIST SET BM_COMMENTS = '" & Left(elpRept(1).Caption, 50) & "' WHERE BM_WRKEMP = '" & glbUserID & "' AND BM_BENE_ID = " & xID & " "
        gdbAdoIhr001.Execute SQLQ
        Data1.Refresh
        Data1.Recordset.Find "BM_BENE_ID = " & xID & " "
    End If
End Sub

Private Sub cmdSele_Click()
Dim xTot As Integer
Dim I As Integer
Dim xList As String

    xTot = lvw.ListItems.count
    
    xSeleTot = 1
    xList = ""
    For I = 1 To xTot
        lvw.ListItems(I).Checked = True
    Next
    
End Sub



Private Sub elpRept_Change(Index As Integer)
    'If Index = 1 Then
    '    Call cmdOne_Click
    'End If
End Sub

Private Sub Form_Load()
Dim rsDiv As New ADODB.Recordset
Dim SQLQ As String
    frmCheckListView.Height = 6690
    
    Screen.MousePointer = DEFAULT
    If glbWFC_IPPopFormName = "WIADivisoinList" Then
        'Me.Caption = glbWFC_IPPopFormName
        Me.Caption = "WAI Divisoin List"
        frmCheckListView.Width = 6720
        frmCheckListView.Height = 5145

        lvw.ColumnHeaders(1).Text = "Select"
        lvw.ColumnHeaders(2).Text = "Division Description"
        lvw.ColumnHeaders(3).Text = "Division Code"
        SQLQ = "SELECT * FROM HR_DIVISION WHERE DV_WIA = 1 ORDER BY Division_Name" ' Not (DV_WIA IS NULL OR DV_WIA = '')"
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsDiv.EOF
            lvw.ListItems.Add
            lvw.ListItems(lvw.ListItems.count).Checked = False ' True
            lvw.ListItems(lvw.ListItems.count).SubItems(1) = rsDiv("Division_Name")
            lvw.ListItems(lvw.ListItems.count).SubItems(2) = rsDiv("DIV")
            rsDiv.MoveNext
        Loop
        rsDiv.Close
        cmdCancel.Left = 2600
        cmdSele.Visible = True
        cmdClear.Visible = True
    End If
    
    'Ticket #29183 Franks 09/13/2016
    If glbWFC_IPPopFormName = "WFCEmpListWithOldPos" Then
        Call PopulateEmpWithOldPositions
    End If
    
    'Ticket #29438 Franks 11/08/2016
    If glbWFC_IPPopFormName = "WFCEmpListWithRepPosBased" Then
        'If glbWFC_IncePlanID = -100 Then 'From Position Master change
            Call PopulateEmpWithReptPositionBased
        'Else
        '    Call PopulateEmpWithReptEmpBased
        'End If
    End If
    
    'Ticket #29438 Franks 11/08/2016
    If glbWFC_IPPopFormName = "WFCEmpListWithRept" Then
        'If glbWFC_IncePlanID = -100 Then 'From Position Master change
        '    Call PopulateEmpWithReptPositionBased
        'Else
            Call PopulateEmpWithReptEmpBased
        'End If
    End If
    
    If glbWFC_IPPopFormName = "WFCEmpListWithRepTerm" Then 'From Termination
        Call PopulateEmpWithReptEmpBased
    End If
    
    If glbWFC_IPPopFormName = "WFCEmpListWithRepTran" Then 'From Transfer Out
        Call PopulateEmpWithReptEmpBased
    End If
    
    If glbWFC_IPPopFormName = "WFCEmpListWithRepTranIn" Then 'From Transfer In
        Call PopulateEmpWithReptTransferIn
    End If
    
    If glbWFC_IPPopFormName = "UpdateBUFin" Then
        Me.Caption = "Update BU Fin for Year"
        Call ScreenUpdateBUFin
        Call INI_Controls(Me)
    End If
    If glbWFC_IPPopFormName = "UpdateCorpFin" Then
        Me.Caption = "Update Corp Fin for Year"
        Call ScreenUpdateCorpFin
    End If
    If glbWFC_IPPopFormName = "UpdateSalesComm" Then
        Me.Caption = "Update Corp Fin for Year"
        Call ScreenUpdateCorpFin
    End If
    
    If glbWFC_IPPopFormName = "UpdateROIC" Then
        Me.Caption = "Update ROIC for Year"
        Call ScreenUpdateROIC
    End If
    
    If glbWFC_IPPopFormName = "ImpCurrency" Then
        glbOnTop = UCase("frmCheckListView")
        Me.Caption = "Import Currency Exchange Table"
        'Me.MDIChild = True
        Call ScreenImpCurrency
    End If
End Sub

Private Sub ScreenUpdateBUFin()
    lvw.Visible = False
    lbltitle(0).Top = 360
    clpCode(0).Top = 360
    lbltitle(1).Top = 360 + 480
    MskTPerc.Top = 360 + 480
    Me.Height = 3225 - 300
    Me.Width = 6705
    
    lbltitle(0).Caption = lStr("Region")
    lbltitle(0).Visible = True
    clpCode(0).Visible = True
    lbltitle(1).Visible = True
    MskTPerc.Visible = True
End Sub

Private Sub ScreenUpdateCorpFin()
    lvw.Visible = False
    lbltitle(1).Top = 360 + 480
    MskTPerc.Top = 360 + 480
    Me.Height = 3225 - 300
    Me.Width = 6705
    
    lbltitle(1).Visible = True
    MskTPerc.Visible = True
End Sub

Private Sub ScreenUpdateROIC()
    lvw.Visible = False
    lbltitle(1).Top = 360 + 480
    MskTPerc.Top = 360 + 480
    Me.Height = 3225 - 300
    Me.Width = 6705
    
    lbltitle(1).Visible = True
    MskTPerc.Visible = True
    MskTPerc.Format = "#,##0.00;(#,##0.00)"
    lbltitle(1).Caption = "Factor"
    
    lblCriteria(7).Top = 360
    MskFiscalYea2.Top = 360
    lblCriteria(7).Visible = True
    MskFiscalYea2.Visible = True
    
End Sub

Private Sub ScreenImpCurrency()
    lvw.Visible = False
    cmdOK.Visible = False
    cmdCancel.Visible = False
    
    fraImpCurrencyExch.Left = 0
    fraImpCurrencyExch.Top = 0
    fraImpCurrencyExch.BorderStyle = 0
    
    Me.Height = 4515
    Me.Width = 9120
    
    Call MonthDescAdd
    
    fraImpCurrencyExch.Visible = True
    
End Sub

Private Sub MonthDescAdd()
ComMTH.AddItem "00-Annual Average Rate"
ComMTH.AddItem "01-January"
ComMTH.AddItem "02-February"
ComMTH.AddItem "03-March"
ComMTH.AddItem "04-April"
ComMTH.AddItem "05-May"
ComMTH.AddItem "06-June"
ComMTH.AddItem "07-July"
ComMTH.AddItem "08-August"
ComMTH.AddItem "09-September"
ComMTH.AddItem "10-October"
ComMTH.AddItem "11-November"
ComMTH.AddItem "12-December"
ComMTH.ListIndex = -1
End Sub

Private Sub mdPrint_Click()

End Sub

Private Sub optAT_Click(Index As Integer)
    If Index = 1 Then
        elpRept(1).LookupType = TERM
    Else
        elpRept(1).LookupType = 0  '0 = ACTIVE. I cannot put as ACTIVE because it's changing to "Active" and that does not switch the lookup to ACTIVE employees
    End If
End Sub

Private Sub txtCDN_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtUS_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub WFCImpCurrencyFunc(xYear, xMTHSEQ, xMTHDesc, xCONVERT_NO, xORDER, xCURRENCYIND1, xCURRENCYIND2, xCOUNTRY, xCURRENCYINDF, xRate)
Dim rsVE As New ADODB.Recordset
Dim SQLQ As String

SQLQ = "SELECT * FROM HRIP_CURRENCY_EXCHG WHERE IP_YEAR = " & xYear & " "
SQLQ = SQLQ & "AND IP_MTH_SEQ = '" & xMTHSEQ & "' "
SQLQ = SQLQ & "AND IP_CONVERT_NO = " & xCONVERT_NO & " "
SQLQ = SQLQ & "AND IP_ORDER = " & xORDER & " "

rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If rsVE.EOF Then
    rsVE.AddNew
    rsVE("IP_YEAR") = xYear
    rsVE("IP_MTH_SEQ") = xMTHSEQ ' Left(ComMTH.Text, 2)
    rsVE("IP_MTH_DESC") = Left(xMTHDesc, 30) 'Left(Mid(ComMTH.Text, 4, 30), 30)
End If

rsVE("IP_CURRENCYIND1") = xCURRENCYIND1 'clpCode(0).Text
If Len(xCURRENCYIND2) > 0 Then
    rsVE("IP_CURRENCYIND2") = xCURRENCYIND2
End If
rsVE("IP_CONVERT_NO") = xCONVERT_NO '1
rsVE("IP_ORDER") = xORDER 'x%
If Len(xCOUNTRY) > 0 Then
    rsVE("IP_COUNTRY") = Left(xCOUNTRY, 10)
End If
rsVE("IP_CURRENCYINDF") = xCURRENCYINDF
If Len(xRate) > 0 Then
    If IsNumeric(xRate) Then rsVE("IP_RATE") = Round(xRate, 5)
End If
rsVE("IP_LDATE") = Date
rsVE("IP_LTIME") = Time$
rsVE("IP_LUSER") = glbUserID
rsVE.Update
rsVE.Close

End Sub

Private Sub PopulateEmpWithReptTransferIn()
Dim rsEmp As New ADODB.Recordset
Dim rsTermEmp As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpName, xTranEmpName
Dim xPosDesc
Dim xTmpCode As String

        cmdOK.Caption = "Update"
        cmdCancel.Caption = "Close"
        cmdView.Visible = True
        cmdPrint.Visible = True
        cmdCancelTran.Left = 3600
        cmdCancelTran.Visible = True
            
        frmAT.Visible = True
        optAT(1).Value = True
        
        frmCheckListView.Width = 8535 + 400 + 2700
        frmCheckListView.Height = 5145 + 500 '- 500

        lblNote1.Left = 100
        lblNote1.Top = 100
        lblNote1.Visible = True

        lvw.Visible = False
        
        Data1.ConnectionString = glbAdoIHRDBW
        
        fraReptEmpLis2.Top = 760
        fraReptEmpLis2.Left = 100
        fraReptEmpLis2.Height = 3735
        fraReptEmpLis2.Visible = True
        
        gdbAdoIhr001W.BeginTrans
        gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
        gdbAdoIhr001W.CommitTrans
        
        SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
        rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
            
        xTmpCode = glbWFCNewPosJob
        xPosDesc = getPosDesc(xTmpCode)
        Me.Caption = "Report To " & glbWFCNewPosJob & " " & xPosDesc & " Employee List" ' "Employee Reporting List"
        xLocRptName = "Report To " & glbWFCNewPosJob & " " & xPosDesc & " Employee List"
        lblNote1.Caption = "The following Employees have their Reporting Authority #1. " & "Who should be their Interim/New Reporting Authority?"
            
        xTranEmpName = ""
        SQLQ = "SELECT * FROM Term_HREMP INNER JOIN Term_HRTRMEMP ON Term_HRTRMEMP.TERM_SEQ = Term_HREMP.TERM_SEQ WHERE ED_EMPNBR = " & glbWFC_IncePlanID & " "
        SQLQ = SQLQ & "AND Term_Reason = 'TOUT' "
        SQLQ = SQLQ & "ORDER BY Term_DOT DESC" '
        rsTermEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTermEmp.EOF Then
            xTranEmpName = rsTermEmp("ED_SURNAME") & ", " & rsTermEmp("ED_FNAME")
        End If
        rsTermEmp.Close
        
        'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 "
        'SQLQ = SQLQ & "AND JH_REPTAU = " & glbWFC_IncePlanID & ") "
        SQLQ = "SELECT HREMP.*,JH_REPTAU FROM HREMP LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE JH_CURRENT = 1 "
        SQLQ = SQLQ & "AND ED_EMPNBR IN (SELECT TT_EMPNBR FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "') "
        SQLQ = SQLQ & "ORDER BY ED_SURNAME, ED_FNAME"

        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsEmp.EOF
            If IsNull(rsEmp("JH_REPTAU")) Then 'Ticket #30465 Franks 08/03/2017
                'Do not add to the list
            Else
                rsBGTMP.AddNew
                rsBGTMP("BM_COMPNO") = "001"
                rsBGTMP("BM_WAITPERIOD") = rsEmp("ED_EMPNBR")
                rsBGTMP("BM_BCODE_DESC") = Left(rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME"), 50)
                rsBGTMP("BM_NUM_1") = rsEmp("JH_REPTAU") 'current
                xEmpName = GetEmpData(rsEmp("JH_REPTAU"), "ED_SURNAME") & ", " & GetEmpData(rsEmp("JH_REPTAU"), "ED_FNAME")
                rsBGTMP("BM_TXT_1") = Left(xEmpName, 50)
                rsBGTMP("BM_WRKID") = glbWFC_IncePlanID 'rsEmp("JH_REPTAU") 'new RA #
                rsBGTMP("BM_COMMENTS") = Left(xTranEmpName, 50) ' Left(xEmpName, 50)
                rsBGTMP("BM_WRKEMP") = glbUserID
                rsBGTMP.Update
            End If
            rsEmp.MoveNext
        Loop
        rsEmp.Close
        
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
        SQLQ = SQLQ & "ORDER BY BM_BCODE_DESC " ' BM_BCODE, BM_ACTION "
        Data1.RecordSource = SQLQ
        Data1.Refresh
        
        Call Display_Value
        
        Call INI_Controls(Me)

End Sub

Private Sub PopulateEmpWithReptPositionBased()
Dim rsEmp As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpName

        cmdOK.Caption = "Update"
        cmdCancel.Caption = "Close"
        cmdView.Visible = True
        cmdPrint.Visible = True
        
        frmCheckListView.Width = 8535 + 400 + 2700
        frmCheckListView.Height = 5145 + 500 '- 500

        lblNote1.Left = 100
        lblNote1.Top = 100
        lblNote1.Visible = True

        lvw.Visible = False
        
        Data1.ConnectionString = glbAdoIHRDBW
        
        fraReptEmpLis2.Top = 760
        fraReptEmpLis2.Left = 100
        fraReptEmpLis2.Height = 3495
        fraReptEmpLis2.Visible = True
        
        gdbAdoIhr001W.BeginTrans
        gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
        gdbAdoIhr001W.CommitTrans
        
        SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
        rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
            

        Me.Caption = "Employee Reporting List"
        xLocRptName = "Employee Reporting List"
        lblNote1.Caption = "The following Employees have their Reporting Authority #1. " & "Who should be their Interim/New Reporting Authority?"
            
        'SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 "
        'SQLQ = SQLQ & "AND JH_REPTAU = " & glbWFC_IncePlanID & ") "
        SQLQ = "SELECT HREMP.*,JH_REPTAU FROM HREMP LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR WHERE JH_CURRENT = 1 "
        SQLQ = SQLQ & "AND JH_REPTAU IN (SELECT TT_EMPNBR FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "') "
        SQLQ = SQLQ & "ORDER BY ED_SURNAME, ED_FNAME"

        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsEmp.EOF
            rsBGTMP.AddNew
            rsBGTMP("BM_COMPNO") = "001"
            rsBGTMP("BM_WAITPERIOD") = rsEmp("ED_EMPNBR")
            rsBGTMP("BM_BCODE_DESC") = Left(rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME"), 50)
            rsBGTMP("BM_NUM_1") = rsEmp("JH_REPTAU") 'current
            xEmpName = GetEmpData(rsEmp("JH_REPTAU"), "ED_SURNAME") & ", " & GetEmpData(rsEmp("JH_REPTAU"), "ED_FNAME")
            rsBGTMP("BM_TXT_1") = Left(xEmpName, 50)
            rsBGTMP("BM_WRKID") = rsEmp("JH_REPTAU") 'new RA #
            rsBGTMP("BM_COMMENTS") = Left(xEmpName, 50)
            rsBGTMP("BM_WRKEMP") = glbUserID
            rsBGTMP.Update
            rsEmp.MoveNext
        Loop
        rsEmp.Close
        
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
        SQLQ = SQLQ & "ORDER BY BM_BCODE_DESC " ' BM_BCODE, BM_ACTION "
        Data1.RecordSource = SQLQ
        Data1.Refresh
        
        Call Display_Value
        
        Call INI_Controls(Me)

End Sub

Private Sub PopulateEmpWithReptEmpBased()
Dim rsEmp As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpName

        cmdOK.Caption = "Update"
        cmdCancel.Caption = "Close"
        cmdView.Visible = True
        cmdPrint.Visible = True
        If glbWFC_IPPopFormName = "WFCEmpListWithRepTerm" Or glbWFC_IPPopFormName = "WFCEmpListWithRepTran" Then
            cmdCancelTran.Left = 3600
            cmdCancelTran.Visible = True
        End If
        
        frmCheckListView.Width = 8535 + 400 + 2700
        frmCheckListView.Height = 5145 + 500 '- 500

        lblNote1.Left = 100
        lblNote1.Top = 100
        lblNote1.Visible = True

        lvw.Visible = False
        
        Data1.ConnectionString = glbAdoIHRDBW
        
        fraReptEmpLis2.Top = 760
        fraReptEmpLis2.Left = 100
        fraReptEmpLis2.Height = 3495
        fraReptEmpLis2.Visible = True
        
        gdbAdoIhr001W.BeginTrans
        gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
        gdbAdoIhr001W.CommitTrans
        
        SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
        rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
            

        xEmpName = GetEmpData(glbWFC_IncePlanID, "ED_SURNAME") & ", " & GetEmpData(glbWFC_IncePlanID, "ED_FNAME")
        Me.Caption = "Employees Reporting to #" & glbWFC_IncePlanID & " " & xEmpName
        xLocRptName = "Report To " & "#" & glbWFC_IncePlanID & " " & xEmpName & " Employee List"
        lblNote1.Caption = "The following employees have #" & glbWFC_IncePlanID & " " & xEmpName & " as their Reporting Authority #1. " & "Who should be their Interim/New Reporting Authority?"
    
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 "
        SQLQ = SQLQ & "AND JH_REPTAU = " & glbWFC_IncePlanID & ") "
        SQLQ = SQLQ & "ORDER BY ED_SURNAME, ED_FNAME"

        
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsEmp.EOF
            rsBGTMP.AddNew
            rsBGTMP("BM_COMPNO") = "001"
            rsBGTMP("BM_WAITPERIOD") = rsEmp("ED_EMPNBR")
            rsBGTMP("BM_BCODE_DESC") = Left(rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME"), 50)
            rsBGTMP("BM_NUM_1") = glbWFC_IncePlanID 'current
            rsBGTMP("BM_TXT_1") = Left(xEmpName, 50)
            If glbWFC_IPPopFormName = "WFCEmpListWithRepTerm" Or glbWFC_IPPopFormName = "WFCEmpListWithRepTran" Then
                'do not populate the new RA on Term or Tansfer
            Else
                rsBGTMP("BM_WRKID") = glbWFC_IncePlanID 'new RA #
                rsBGTMP("BM_COMMENTS") = Left(xEmpName, 50)
            End If
            rsBGTMP("BM_WRKEMP") = glbUserID
            rsBGTMP.Update
            rsEmp.MoveNext
        Loop
        rsEmp.Close
        
        SQLQ = "SELECT * FROM HRBENGRPLIST "
        SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
        SQLQ = SQLQ & "ORDER BY BM_BCODE_DESC " ' BM_BCODE, BM_ACTION "
        Data1.RecordSource = SQLQ
        Data1.Refresh
        
        Call Display_Value
        
        Call INI_Controls(Me)
End Sub

Private Sub PopulateEmpWithOldPositions() 'Ticket #29438 Franks 11/08/2016
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
        Me.Caption = "Employees with old Position"
        cmdOK.Caption = "Update"
        cmdCancel.Caption = "Close"
        cmdSele.Visible = True
        cmdClear.Visible = True
        
        frmCheckListView.Width = 6720
        frmCheckListView.Height = 5145

        lvw.ColumnHeaders(1).Text = "Select"
        lvw.ColumnHeaders(2).Text = "Employee #"
        lvw.ColumnHeaders(3).Text = "Employee Name"
        lvw.ColumnHeaders(1).Width = 800
        lvw.ColumnHeaders(2).Width = 1400
        lvw.ColumnHeaders(3).Width = 4000
        
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 "
        SQLQ = SQLQ & "AND JH_JOB = '" & glbChgTermReason & "') "
        SQLQ = SQLQ & "ORDER BY ED_SURNAME, ED_FNAME"
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsEmp.EOF
            lvw.ListItems.Add
            lvw.ListItems(lvw.ListItems.count).Checked = True
            lvw.ListItems(lvw.ListItems.count).SubItems(1) = rsEmp("ED_EMPNBR")
            lvw.ListItems(lvw.ListItems.count).SubItems(2) = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
            rsEmp.MoveNext
        Loop
        rsEmp.Close
End Sub

Private Sub WFCEmpReptUptFunc()
Dim rsEmpPos As New ADODB.Recordset
Dim rsAdd As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim xEmpNo, xPlant
Dim SQLQ As String
Dim I As Integer
Dim xTot
Dim xEDate

Screen.MousePointer = HOURGLASS

    If IsDate(lblStDate.Caption) Then
        xEDate = lblStDate.Caption
    Else
        xEDate = Date
    End If

    'Ticket #29484 Franks 11/22/2016

    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND NOT (BM_NUM_1 = BM_WRKID) "
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rs.EOF
        xEmpNo = rs("BM_WAITPERIOD")
        xOldReptNo = rs("BM_NUM_1")
        xNewReptNo = rs("BM_WRKID")
        Call WFCPosReptsUpd(xEmpNo, xOldReptNo, xNewReptNo, xEDate)
        rs.MoveNext
    Loop
    rs.Close
    

Screen.MousePointer = DEFAULT

End Sub

Sub Display_Value()
    If Not Data1.Recordset.EOF And Not Data1.Recordset.BOF Then
        If Not IsNull(Data1.Recordset("BM_WRKID")) Then
            elpRept(1).Text = Data1.Recordset("BM_WRKID")
        Else
            elpRept(1).Text = ""
        End If
    End If
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Function IsNewRASameAsOldRA()
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND BM_WRKID = BM_NUM_1 "
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        retVal = True
    End If
    rs.Close
    IsNewRASameAsOldRA = retVal
    
End Function

Private Function IsNewRABlank()
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    SQLQ = SQLQ & "AND BM_WRKID IS NULL "
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        retVal = True
    End If
    rs.Close
    IsNewRABlank = retVal
End Function

Private Function IsRAChanged()
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rs.EOF
        If Not (rs("BM_NUM_1") = rs("BM_WRKID")) Then
            retVal = True
        End If
        rs.MoveNext
    Loop
    rs.Close
    IsRAChanged = retVal
End Function

Public Sub cmdView_Click()
Dim RHeading As String

''Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
''This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
'Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
'
'RHeading = Me.Caption
'Me.vbxCrystal.WindowTitle = RHeading & " Report"
'Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Destination = 0
'Me.vbxCrystal.Action = 1

Me.vbxCrystal.Reset
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup
Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "rgRepEmpList.rpt"
'Me.vbxCrystal.Formulas(1) = "rptTitle = 'Report To " & xLocRptName & " Employee List'"
Me.vbxCrystal.Formulas(1) = "rptTitle = '" & xLocRptName & "'"
Me.vbxCrystal.WindowTitle = RHeading & " Report To List"
Me.vbxCrystal.Connect = RptODBC_SQL
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub
Public Sub cmdPrint_Click()
'Dim RHeading As String
'
'RHeading = Me.Caption
'Me.vbxCrystal.WindowTitle = RHeading & " Report"
'Me.vbxCrystal.BoundReportHeading = RHeading
'Me.vbxCrystal.Destination = 1
'Me.vbxCrystal.Action = 1
    Call cmdView_Click
End Sub

