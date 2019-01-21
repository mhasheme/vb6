VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEMAIL 
   Caption         =   "Email Setup"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   10170
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog AttachmentDialog 
      Left            =   3600
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSPanel SSPanel1 
      Align           =   1  'Align Top
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Tag             =   "Get SMTP Information from Company Preference"
      ToolTipText     =   "Get SMTP Information from Company Preference"
      Top             =   0
      Width           =   10170
      _Version        =   65536
      _ExtentX        =   17939
      _ExtentY        =   13996
      _StockProps     =   15
      ForeColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      FloodColor      =   16744576
      Begin VB.CommandButton cmdUpdSMTP 
         Appearance      =   0  'Flat
         Caption         =   "From Company Preference..."
         Height          =   285
         Left            =   4680
         TabIndex        =   30
         Tag             =   "Get SMTP info. from Company Preference"
         ToolTipText     =   "Get SMTP info. from Company Preference"
         Top             =   4246
         Width           =   2295
      End
      Begin VB.CommandButton cmdUpdALLSMTP 
         Appearance      =   0  'Flat
         Caption         =   "Update ALL SMTP Info. from Company Preference"
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
         Left            =   240
         TabIndex        =   29
         Tag             =   "Update All User's SMTP info. from Company Preference"
         ToolTipText     =   "Update All User's SMTP info. from Company Preference"
         Top             =   6120
         Width           =   4575
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         DataField       =   "EM_PORT"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         TabIndex        =   27
         Tag             =   "SMTP Port"
         Top             =   5350
         Width           =   2475
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         DataField       =   "EM_PASSWORD"
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2100
         PasswordChar    =   "*"
         TabIndex        =   26
         Tag             =   "SMTP Password"
         Top             =   4982
         Width           =   2475
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         DataField       =   "EM_USERNAME"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         TabIndex        =   25
         Tag             =   "SMTP Username"
         Top             =   4614
         Width           =   2475
      End
      Begin VB.TextBox txtServer 
         Appearance      =   0  'Flat
         DataField       =   "EM_SERVER"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         TabIndex        =   24
         Tag             =   "SMTP Server"
         Top             =   4246
         Width           =   2475
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         DataField       =   "EM_ADDRESS"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2100
         TabIndex        =   23
         Tag             =   "Email Address"
         Top             =   3878
         Width           =   2475
      End
      Begin VB.CommandButton cmdImport 
         Appearance      =   0  'Flat
         Caption         =   "Import"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   9000
         TabIndex        =   20
         Tag             =   "Import the File"
         Top             =   6480
         Visible         =   0   'False
         Width           =   855
      End
      Begin Crystal.CrystalReport crwMain 
         Left            =   2880
         Top             =   7440
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         ReportSource    =   1
         BoundReportHeading=   "INFO:HR Email Setup"
         PrintFileLinesPerPage=   60
         GridSource      =   "vbxTrueGrid"
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.CommandButton cmdImportFile 
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
         Height          =   280
         Left            =   8520
         TabIndex        =   6
         Tag             =   "Select File to Import"
         Top             =   6480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   150
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "00-File Name to import"
         Top             =   6480
         Visible         =   0   'False
         Width           =   6375
      End
      Begin VB.CheckBox chkSend 
         Alignment       =   1  'Right Justify
         Caption         =   "Send Email"
         DataField       =   "EM_SEND"
         Height          =   255
         Left            =   8550
         TabIndex        =   4
         Top             =   4020
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CheckBox chkReceive 
         Alignment       =   1  'Right Justify
         Caption         =   "Receive Email"
         DataField       =   "EM_RECEIVE"
         Height          =   255
         Left            =   8550
         TabIndex        =   3
         Top             =   3750
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.CheckBox chkIsSuper 
         Alignment       =   1  'Right Justify
         Caption         =   "Supervisor"
         DataField       =   "EM_IS_SUPER"
         Height          =   255
         Left            =   8550
         TabIndex        =   2
         Top             =   3480
         Width           =   1485
      End
      Begin VB.CommandButton cmdIP 
         Appearance      =   0  'Flat
         Caption         =   "Send Email"
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Tag             =   "Recalculate for all employees"
         Top             =   4920
         Visible         =   0   'False
         Width           =   2295
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "fvemail.frx":0000
         Height          =   2730
         Left            =   240
         Negotiate       =   -1  'True
         OleObjectBlob   =   "fvemail.frx":0014
         TabIndex        =   1
         Tag             =   "Listing of Emails"
         Top             =   600
         Width           =   9855
      End
      Begin Threed.SSPanel panEEDESC 
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   18120
         _Version        =   65536
         _ExtentX        =   31962
         _ExtentY        =   873
         _StockProps     =   15
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   0
         BevelInner      =   2
         Font3D          =   1
         Begin VB.CommandButton cmdFindUser 
            Caption         =   "Find User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   9000
            TabIndex        =   28
            Top             =   90
            Width           =   1095
         End
         Begin Threed.SSPanel Panel3D2 
            Height          =   1170
            Left            =   0
            TabIndex        =   8
            Top             =   525
            Width           =   1275
            _Version        =   65536
            _ExtentX        =   2646
            _ExtentY        =   1323
            _StockProps     =   15
            Caption         =   "Panel3D2"
            ForeColor       =   0
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   1
            Alignment       =   1
         End
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
            Left            =   6840
            TabIndex        =   17
            Top             =   120
            Width           =   1305
         End
         Begin VB.Label lblEEName 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Name"
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
            Left            =   3045
            TabIndex        =   16
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lblEENum 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            Caption         =   "Employee#"
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
            Left            =   1320
            TabIndex        =   10
            Top             =   120
            Width           =   1185
         End
         Begin VB.Label lblEENumber 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "User ID"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   150
            Width           =   540
         End
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   375
         Left            =   720
         Top             =   7320
         Visible         =   0   'False
         Width           =   2115
         _ExtentX        =   3731
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
         Caption         =   "HR_EMAIL"
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
      Begin INFOHR_Controls.EmployeeLookup elpEmpNbr 
         DataField       =   "EM_USERID"
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Tag             =   "User ID"
         Top             =   3510
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   503
         ShowUnassigned  =   1
         RefreshDescriptionWhen=   2
      End
      Begin VB.Label lblPort 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Port:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   21
         Top             =   5395
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   6525
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgHelp 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   285
         Picture         =   "fvemail.frx":B904
         Stretch         =   -1  'True
         Top             =   6495
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Password:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   5003
         Width           =   1290
      End
      Begin VB.Label lblServer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Server:"
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
         Height          =   255
         Left            =   300
         TabIndex        =   14
         Top             =   4279
         Width           =   1605
      End
      Begin VB.Label lblAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address:"
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
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   3887
         Width           =   1605
      End
      Begin VB.Label lblEmpNbr 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID:"
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
         Height          =   195
         Left            =   300
         TabIndex        =   12
         Top             =   3555
         Width           =   1605
      End
      Begin VB.Label lblUsername 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Username:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   4671
         Width           =   1290
      End
   End
End
Attribute VB_Name = "frmEMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim Unloading As Boolean        ' True if DB is not set up, and we are unloading form
                                ' This prevents multiple errors from occuring.
Dim RSDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim xImportFile As String

' Lock or unlock the form for user entry.  Pass True to lock controls.
Private Sub LockForm(Locked As Boolean)
    elpEmpNbr.Enabled = Not Locked
    txtAddress.Enabled = Not Locked
    txtServer.Enabled = Not Locked
    txtUsername.Enabled = Not Locked
    txtPassword.Enabled = Not Locked
    txtPort.Enabled = Not Locked
    
    'Ticket #26529 - SMTP information available on Company Preference.
    If Not Locked Then
        If gsSMTPINFO Then
            cmdUpdSMTP.Enabled = True
            If Not fglbNew Then
                cmdUpdALLSMTP.Enabled = True
            Else
                cmdUpdALLSMTP.Enabled = False
            End If
        Else
            cmdUpdSMTP.Enabled = False
            cmdUpdALLSMTP.Enabled = False
        End If
    End If
    
End Sub

Sub cmdCancel_Click()
    fglbNew = False
    Data1.Refresh
    Call Display_Value
    
    panEEDESC.Enabled = True
End Sub

Sub cmdClose_Click()
    Unload Me
End Sub

Sub cmdDelete_Click()
Dim Msg, a%

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
RSDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh

If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False

Call SET_UP_MODE

End Sub

Sub cmdNew_Click()
   ' Data1.Recordset.AddNew
    Call Set_Control("B", Me)
    
    fglbNew = True
    chkReceive.Value = 1
    chkSend.Value = 1
    Call SET_UP_MODE
    'If Not glbSQL And Not glbOracle Then 'Added by Franks Mar 4,03 Ticket 3758
        txtServer = "xxx.yyy"
    'End If
    
    panEEDESC.Enabled = False
End Sub

Sub cmdOK_Click()
Dim xID
Dim rsDup As ADODB.Recordset
  xID = elpEmpNbr
    If elpEmpNbr.Caption = "Unassigned" Then
        MsgBox "User ID must be valid.", vbExclamation + vbOKOnly, "Invalid User ID"
        elpEmpNbr.SetFocus
        Exit Sub
    End If
    If Len(elpEmpNbr.Text) = 0 Then
        MsgBox "User ID is a required field.", vbExclamation + vbOKOnly, "Missing User ID"
        Exit Sub
    End If
    If Len(txtAddress.Text) = 0 Then
        MsgBox "Email Address is a required field.", vbExclamation + vbOKOnly, "Missing Email Address"
        txtAddress.SetFocus
        Exit Sub
    End If
    If fglbNew Then
        Set rsDup = Data1.Recordset.Clone
        If rsDup.RecordCount > 0 Then
            rsDup.MoveFirst
            rsDup.Find "EM_USERID='" & Replace(xID, "'", "''") & "'"
            If Not rsDup.EOF Then
                MsgBox "Duplicate User ID."
                elpEmpNbr.SetFocus
                Exit Sub
            End If
        End If
    End If
    
    'If Not glbSQL And Not glbOracle Then 'Added by Franks Mar 4,03 Ticket 3758
        If Len(txtServer.Text) = 0 Then
            MsgBox "SMTP Server is a required field.", vbExclamation + vbOKOnly, "Missing SMTP Server"
            Exit Sub
        End If
        If Not IsValidDNSName(txtServer.Text) Then
            MsgBox "SMTP Server must be in xxx.yyy format.", vbExclamation + vbOKOnly, "Invalid SMTP Server"
            Exit Sub
        End If
    'End If
    If Not IsEmail(txtAddress.Text) Then
        MsgBox "Email address must be in xxx@yyy.zzz format.", vbExclamation + vbOKOnly, "Invalid Email Address"
        Exit Sub
    End If
    If Not InSecure(elpEmpNbr.Text) Then
        MsgBox "Employee must be set up to log in to info:HR (listed in the Security Master) before setting up email.", vbExclamation + vbOKOnly, "Employee Not in SECURE"
        Exit Sub
    End If
    If Len(txtPort.Text) > 0 Then
        If Not IsNumeric(txtPort.Text) Then
            MsgBox "SMTP Port must be numeric.", vbExclamation + vbOKOnly, "Invalid SMTP Port"
            Exit Sub
        End If
    End If
    
    If fglbNew Then RSDATA.AddNew
    If txtUsername.Text = "" Then
        txtUsername.DataChanged = False
        RSDATA!EM_USERNAME = Null
    End If
    If txtPassword.Text = "" Then
        txtPassword.DataChanged = False
         RSDATA!EM_PASSWORD = Null
    End If
    RSDATA!EM_LDATE = Date
    RSDATA!EM_LTIME = Time$
    RSDATA!EM_LUSER = glbUserID
    
    gdbAdoIhr001.BeginTrans
    Call Set_Control("U", Me, RSDATA)
    
    RSDATA.Update
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
    Data1.Recordset.Find "EM_USERID='" & Replace(xID, "'", "''") & "'"
    fglbNew = False
    
    Call Display_Value

    panEEDESC.Enabled = True
'Call SET_UP_MODE
End Sub

Sub cmdPrint_Click()
    crwMain.Destination = 1
    crwMain.Action = 1
End Sub

Sub cmdView_Click()
    'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
    'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
    Me.crwMain.WindowShowPrintSetupBtn = True
    
    crwMain.Destination = 0
    crwMain.Action = 1
End Sub

Private Sub cmdFindUser_Click()
    Dim SaveID, SaveName, xTxtEEID, xLblEEName

    SaveID = glbLUserID
    SaveName = glbLUserNAME
    
    frmUFIND.Show 1
    
    If glbEEOK Then
        'txtUSERID = glbLUserID
        'txtEEName = glbLUserNAME
        'glbLUserNAME = SaveName
        'glbLUserID = SaveID

        If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
            glbSecUSERID = glbLUserID
                        
            Data1.Recordset.Requery
            Data1.Recordset.Find "EM_USERID='" & Replace(glbLUserID, "'", "''") & "'"
            
            ''????Ticket #24808 - If the User is Templated based then retrieve the Template Name of this User to retrieve
            ''Template's Profile instead of User's Profile. If the User's Security is not based on Template or is TEMPLATE then
            ''retrieve the respective User's record
            'If cmbSecTemplate = "" Or cmbSecTemplate = "TEMPLATE" Then
            '    'User is normal user or is Template itself
            '    glbSecUSERID = glbLUserID
            'Else
            '    'User's Profile is based on Template
            '    glbSecUSERID = cmbSecTemplate
            'End If
            
            'Call Display_Values
            
            ''????Ticket #24808 - Reset this global variable back to User ID
            'glbSecUSERID = glbLUserID

        End If
    End If
End Sub

Private Sub cmdImport_Click()
    Dim DgDef, Title$, Msg$, Response%
    
    If Trim(txtFileName.Text) = "" Then
        MsgBox "File to import not selected. Please select the file to import.", vbExclamation
        cmdImportFile.SetFocus
        Exit Sub
    ElseIf Dir(txtFileName.Text) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & txtFileName.Text & "]", vbExclamation
        cmdImportFile.SetFocus
        Exit Sub
    Else
        Title$ = "Email Setup Import"
        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
        Msg$ = "Are you sure you want to import this Email Setup file?"
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Exit Sub
        End If
        
        Call Load_EmailSetup
        
        Data1.Refresh
    End If
End Sub

Private Sub cmdImport_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdImportFile_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdIP_Click()
    Dim MailBody As String
    Dim LocCode As String, LocDesc As String
    
    On Error GoTo ErrorHandler
    glbUserID = elpEmpNbr
    glbWFCEmailTest = True
    If glbBurlTech Then
        If Len(txtServer.Text) > 0 Then
            glbSMTPServerIP = txtServer.Text
        End If
    End If
    Load frmSendEmail
    If glbWFC Then
        frmSendEmail.txtSubject.Text = "info:HR Termination Notice"
        frmSendEmail.txtTo.Text = txtAddress.Text
        MailBody = "This email is for test only " & vbCrLf & vbCrLf
        MailBody = MailBody & "By David Hili " & vbCrLf
    End If
    If glbBurlTech Then
        frmSendEmail.txtSubject.Text = "info:HR Counselling Notice"
        frmSendEmail.txtTo.Text = txtAddress.Text
        MailBody = "This email is for test only " & vbCrLf & vbCrLf
        MailBody = MailBody & "By Chris Watts" & vbCrLf
    End If
    MailBody = MailBody & "Date: " & CVDate(Now) & vbCrLf & vbCrLf
    frmSendEmail.txtBody.Text = MailBody
    
    frmSendEmail.Show 1
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
End Sub

Private Sub cmdUpdALLSMTP_Click()
    'Ticket #26529 - Get SMTP information from Company Preference.
    If gsSMTPINFO Then
        Dim xSMTPInfo
        Dim SQLQ As String
        Dim Title$, Msg$, DgDef As Variant, Response%
        
        'Are you sure you want to update all User's information with SMTP information from Company Preference screen?
        Title = "Update ALL User's SMTP Information"
        Msg$ = "Are you sure you want to update ALL User's SMTP information with SMTP information from Company Preference screen?"
        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
        Response = MsgBox(Msg, DgDef, Title)    ' Get user response.
        If Response = IDNO Then
            Exit Sub
        End If
        
        xSMTPInfo = Split(GetComPreferSMTP("SMTP_INFORMATION"), "|")
        If Len(xSMTPInfo(3)) = 0 Then
            SQLQ = "UPDATE HR_EMAIL SET EM_SERVER = '" & xSMTPInfo(0) & "' , EM_USERNAME = '" & xSMTPInfo(1) & "', EM_PASSWORD = '" & xSMTPInfo(2) & "'"  'WHERE HP_FUN_NAME = 'SMTP_INFORMATION'"
        Else
            SQLQ = "UPDATE HR_EMAIL SET EM_SERVER = '" & xSMTPInfo(0) & "' , EM_USERNAME = '" & xSMTPInfo(1) & "', EM_PASSWORD = '" & xSMTPInfo(2) & "' ,EM_PORT = " & xSMTPInfo(3) '& " WHERE HP_FUN_NAME = 'SMTP_INFORMATION'"
        End If
        gdbAdoIhr001.Execute SQLQ
        Data1.Refresh
        Call Display_Value
    End If

End Sub

Private Sub cmdUpdALLSMTP_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdUpdSMTP_Click()
    'Ticket #26529 - Get SMTP information from Company Preference.
    If gsSMTPINFO Then
        Dim xSMTPInfo
        
        xSMTPInfo = Split(GetComPreferSMTP("SMTP_INFORMATION"), "|")
        txtServer = xSMTPInfo(0)
        txtUsername = xSMTPInfo(1)
        txtPassword = xSMTPInfo(2)
        txtPort = xSMTPInfo(3)
    End If
    
End Sub

Private Sub cmdUpdSMTP_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    fCancelDisplay = True
    If ErrorNumber = 3639 Then
        MsgBox "Database not set up for email capability.  Please call HR Systems Support for details on enabling this.", vbCritical + vbOKOnly, "Missing Table"
        Unloading = True
    Else
        MsgBox "Error occured in ADO Data Control." & vbCrLf & Description, vbCritical + vbOKOnly, "Error #" & ErrorNumber
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = Me.name
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

glbOnTop = Me.name

    Call INI_Controls(Me)
    
    If glbWFC Then 'For Test
        If glbUserID = "3142" Or glbUserID = "3517" Or glbUserID = "3079" Then
            cmdIP.Visible = True
        End If
    End If
    
    If glbBurlTech Then
        cmdIP.Visible = True
    End If
    
    If glbOttawaCCAC Then
        chkReceive.Visible = True
    '    chkSend.Visible = True
    End If
    elpEmpNbr.LookupType = 2

    On Error GoTo ErrorHandler
    
    'If glbSQL Or glbOracle Then
    '    lblServer.Font.Bold = False
    'End If
    
    Data1.ConnectionString = glbAdoIHRDB
    Data1.RecordSource = "SELECT * FROM HR_EMAIL ORDER BY EM_USERID"
    Data1.Refresh
        
    Exit Sub
    
ErrorHandler:
    If Unloading = False Then
        MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
        Unload Me
    End If
End Sub

Private Function GetEmpName(UserID As String, EmpLabel As Label)
    Dim rsEmp As New ADODB.Recordset
    On Error GoTo EE_Err
    
    EmpLabel.Visible = False
    EmpLabel.Caption = "Unassigned"
    If Len(UserID) > 0 Then
        EmpLabel.Visible = True
        rsEmp.Open "SELECT USERNAME FROM HR_SECURE_BASIC WHERE USERID = '" & Replace(UserID, "'", "''") & "'", gdbAdoIhr001
        If Not rsEmp.EOF Then
            EmpLabel.Caption = rsEmp("USERNAME")
        End If
        rsEmp.Close
    End If
    Exit Function
    
EE_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Movenext", "HREMP", "FIND")
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub imgHelp_Click()
Dim MsgStr As String
    MsgStr = "Import File must be an Excel Spreadsheet with the following format: "
    MsgStr = MsgStr & Chr(10) & "        1. First row is a Header row."
    MsgStr = MsgStr & Chr(10) & "        2. Data to import must start from 2nd row."
    MsgStr = MsgStr & Chr(10) & "        3. Column order to Import:"
    MsgStr = MsgStr & Chr(10) & vbTab & "a. Column 1: Employee #"
    MsgStr = MsgStr & Chr(10) & vbTab & "b. Column 2: Email Address"
    MsgStr = MsgStr & Chr(10) & vbTab & "c. Column 3: SMTP Server"
    MsgStr = MsgStr & Chr(10) & vbTab & "d. Column 4: SMTP Username"
    MsgStr = MsgStr & Chr(10) & vbTab & "e. Column 5: SMTP Password"
    MsgStr = MsgStr & Chr(10) & vbTab & "f. Column 6: Supervisor ('Y' for Yes or 'N' for No)"
    MsgBox MsgStr, vbInformation, "info:HR - Import File Format"
End Sub

Private Sub Label2_Click()

End Sub

Private Sub txtAddress_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub elpEmpNbr_Change()
    GetEmpName elpEmpNbr.Text, lblEEName
    lblEENum.Caption = elpEmpNbr.Text
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End Sub

Private Function IsEmail(Address As String) As Boolean
    IsEmail = True
    ' Make sure there's an @ in the address
    If InStr(Address, "@") = 0 Then IsEmail = False: Exit Function
    ' Make sure they have at least one period after the @
    If InStr(InStr(Address, "@"), Address, ".") = 0 Then IsEmail = False: Exit Function
    ' Make sure they have text before the period
    If Mid(Address, InStr(Address, "@") + 1, 1) = "." Then IsEmail = False: Exit Function
    ' Make sure they have text after the period
    If Right(Address, 1) = "." Then IsEmail = False: Exit Function
End Function

Private Function IsValidDNSName(Address As String) As Boolean
    IsValidDNSName = True
    ' Make sure they have at least one period
    If InStr(Address, ".") = 0 Then IsValidDNSName = False: Exit Function
    ' Make sure they have text before the period
    If Left(Address, 1) = "." Then IsValidDNSName = False: Exit Function
    ' Make sure they have text after the period
    If Right(Address, 1) = "." Then IsValidDNSName = False: Exit Function
End Function

Private Function InSecure(UserID As String) As Boolean
    Dim rsSecure As New ADODB.Recordset
    On Error GoTo EE_Err
    
    rsSecure.Open "SELECT USERID FROM HR_SECURE_BASIC WHERE USERID = '" & Replace(UserID, "'", "''") & "'", gdbAdoIhr001
    If rsSecure.EOF Then
        InSecure = False
    Else
        InSecure = True
    End If
    rsSecure.Close
    Exit Function
    
EE_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Movenext", "HREMP", "FIND")
End Function
'Private Sub txtEmpNbr_DblClick()
'Dim SaveID, SaveName, xTxtEEID, xLblEEName
'    SaveID = glbLUserID
'    SaveName = glbLUserNAME
'    frmUFIND.Show 1
'    If glbEEOK Then
'        xTxtEEID = glbLUserID
'        xLblEEName = glbLUserNAME
'    Else
'        xTxtEEID = ""
'        xLblEEName = "Unassigned"
'    End If
'    glbLUserNAME = SaveName
'    glbLUserID = SaveID
'    txtEmpNbr.Text = xTxtEEID
'    lblEEName.Caption = xLblEEName
'End Sub
'Private Sub txtEmpNbr_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Private Sub txtFileName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPassword_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtPort_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtServer_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtUsername_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

''' Sam add July 2002 * Remove Binding Control
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        If glbtermopen Then
            RSDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        Call SET_UP_MODE
        Exit Sub
    End If
    

    SQLQ = " SELECT * FROM HR_EMAIL "
    SQLQ = SQLQ & " where EM_USERID = '" & Replace(Data1.Recordset!EM_USERID, "'", "''") & "'"
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic


    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)

Call SET_UP_MODE

End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "SELECT * FROM HR_EMAIL "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
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
RelateMode = RelateSetUp
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Email_Setup    'gSec_Upd_Security Or gSec_Upd_Quick_ESS - 7.9 Enhancement
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
    TF = False
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = True
Else
    UpdateState = OPENING
    TF = False
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = True
LockForm (TF)
End Sub

Private Sub Load_EmailSetup()
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim rsEmp As New ADODB.Recordset
    Dim rsEmail As New ADODB.Recordset
    Dim rsSecure As New ADODB.Recordset
    Dim xSkipped As String
    Dim SQLQ As String
    Dim xEmail, xServer, xUserName, xPassword, xSup As String
    Dim xSuper As Integer
    Dim xNum As Integer
    Dim xRows As Long
    Dim xRow As Long
    Dim xEMPNBR
    
    
    On Error GoTo EmailSetup_Err

    Screen.MousePointer = vbHourglass
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"

    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(txtFileName.Text)
    Set exSheet = exBook.Worksheets(1)
'    xCols = 1
    xSkipped = ""
    xNum = 0
'    ReDim xTitle(xCols)
'    For X = 1 To xCols
'        xTitle(X) = exSheet.Cells(1, X)
'        Debug.Print "case """ & xTitle(X) & """"
'    Next

    xRows = getRows(exSheet)

    For xRow = 2 To xRows
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
     
        xEMPNBR = exSheet.Cells(xRow, 1)
        xEmail = exSheet.Cells(xRow, 2)
        xServer = exSheet.Cells(xRow, 3)
        xUserName = exSheet.Cells(xRow, 4)
        xPassword = exSheet.Cells(xRow, 5)
        xSup = exSheet.Cells(xRow, 6)
        
        If Len(xSup) > 0 Then
            If UCase(xSup) = "Y" Then
                xSuper = 1
            Else
                xSuper = 0
            End If
        Else
            xSuper = 0
        End If

        If Not IsNumeric(xEMPNBR) Or xEMPNBR = 0 Or Trim(xEmail) = "" Or Trim(xServer) = "" Then
            xSkipped = xSkipped & xEMPNBR & "; "
            xNum = xNum + 1
            If xNum = 10 Then
                xSkipped = xSkipped & vbCrLf
                xNum = 0
            End If
        Else
            'Get the User ID
            SQLQ = "SELECT USERID, EMPNBR FROM HR_SECURE_BASIC WHERE EMPNBR = " & xEMPNBR
            rsSecure.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            If Not rsSecure.EOF Then
                If Not IsNull(rsSecure("USERID")) Or rsSecure("USERID") <> "" Then
                    rsEmail.Open "SELECT EM_USERID,EM_ADDRESS,EM_SERVER,EM_USERNAME,EM_PASSWORD,EM_IS_SUPER FROM HR_EMAIL WHERE EM_USERID ='" & Replace(rsSecure("USERID"), "'", "''") & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsEmail.EOF Then
                        SQLQ = "INSERT INTO HR_EMAIL(EM_USERID,EM_ADDRESS,EM_SERVER,EM_USERNAME,EM_PASSWORD,EM_IS_SUPER) VALUES ('" & Replace(rsSecure("USERID"), "'", "''") & "', '" & xEmail & "', '" & xServer & "', '" & xUserName & "', '" & xPassword & "'," & xSuper & ")"
                        gdbAdoIhr001.Execute SQLQ
                    Else
                        rsEmail("EM_ADDRESS") = xEmail
                        rsEmail("EM_SERVER") = xServer
                        rsEmail("EM_USERNAME") = xUserName
                        rsEmail("EM_PASSWORD") = xPassword
                        rsEmail("EM_IS_SUPER") = xSuper
                        rsEmail.Update
                    End If
                    rsEmail.Close
                    Set rsEmail = Nothing
                End If
            Else
                xSkipped = xSkipped & xEMPNBR & "; "
                xNum = xNum + 1
                If xNum = 10 Then
                    xSkipped = xSkipped & vbCrLf
                    xNum = 0
                End If
            End If
            rsSecure.Close
            Set rsSecure = Nothing
        End If
    Next
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    MDIMain.panHelp(0).FloodPercent = 0
    'MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    

    Screen.MousePointer = vbDefault

    If Len(xSkipped) > 0 Then
        MsgBox "The Email address for the following Employee(s) have been skipped:" & vbCrLf & xSkipped, vbOKOnly + vbInformation, "Import Email Addresses"
    Else
        MsgBox "Employee's Email Addresses have been loaded successfully on Email Setup screen.", vbOKOnly + vbInformation, "Import Email Addresses"
    End If

Exit Sub

EmailSetup_Err:
    Set exSheet = Nothing
    Set exBook = Nothing
    exApp.Quit
    Set exApp = Nothing
    
    'MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = vbDefault

    If Err.Number = 1004 Then
        MsgBox "Import file not found, try again.", vbOKOnly + vbExclamation, "Email List File Missing"
        Exit Sub
    Else
        MsgBox Err.Description
        Exit Sub
    End If
End Sub

Private Sub cmdImportFile_Click()
    glbDocName = "EmailSetup"
    
    xImportFile = ""
    AttachmentDialog.DialogTitle = "Select the file to import..."
    AttachmentDialog.Filter = "*.xls;*.xlsx|*.xls;*.xlsx"    '"Word Documents (*.doc;*.docx)|*.doc;*.docx"
    AttachmentDialog.FilterIndex = 1
    AttachmentDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
    AttachmentDialog.ShowOpen
    If Len(AttachmentDialog.FileName) <> 0 Then
        txtFileName.Text = AttachmentDialog.FileName
    Else
        glbDocName = ""
    End If

End Sub

Private Function getRows(exSheet As Object)
Dim X
X = 1
Do While True
    If exSheet.Cells(X, 1) = "" Then
        Exit Do
    Else
        X = X + 1
    End If
Loop
getRows = X - 1
End Function
