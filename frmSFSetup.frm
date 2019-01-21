VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSFSetup 
   Caption         =   "FTP Setup"
   ClientHeight    =   9315
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10485
   DrawStyle       =   1  'Dash
   Icon            =   "frmSFSetup.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9315
   ScaleWidth      =   10485
   WindowState     =   2  'Maximized
   Begin VB.Frame fraXMLLocation 
      Height          =   3975
      Left            =   120
      TabIndex        =   17
      Top             =   3960
      Visible         =   0   'False
      Width           =   9735
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1920
         TabIndex        =   20
         Tag             =   "Disk Drive"
         Top             =   675
         Width           =   3372
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00FFFFFF&
         Height          =   2790
         Left            =   1920
         TabIndex        =   19
         Tag             =   "Path"
         Top             =   1035
         Width           =   3372
      End
      Begin VB.TextBox txtPDFilename 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1920
         MaxLength       =   200
         TabIndex        =   18
         Tag             =   "00-File Name"
         Top             =   240
         Width           =   6975
      End
      Begin VB.Label lblPath 
         BackStyle       =   0  'Transparent
         Caption         =   "Path"
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   735
         Width           =   1500
      End
      Begin VB.Label lblPDFilename 
         BackStyle       =   0  'Transparent
         Caption         =   "XML File Location"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   300
         Width           =   2460
      End
   End
   Begin VB.Frame fraFTP 
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   9615
      Begin VB.TextBox txtHost 
         Appearance      =   0  'Flat
         DataField       =   "HOST"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   2280
         Width           =   2355
      End
      Begin VB.TextBox txtUsername 
         Appearance      =   0  'Flat
         DataField       =   "USERNAME"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Top             =   2700
         Width           =   2355
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         DataField       =   "PASSWORD"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   3090
         Width           =   2355
      End
      Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
         Bindings        =   "frmSFSetup.frx":0442
         Height          =   1845
         Left            =   120
         OleObjectBlob   =   "frmSFSetup.frx":0456
         TabIndex        =   11
         Top             =   240
         Width           =   9135
      End
      Begin Threed.SSCheck chkCurrent 
         DataField       =   "CURRENT_FTP"
         DataSource      =   "Data1"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   450
         _StockProps     =   78
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
      End
      Begin VB.Label lblUsername 
         Caption         =   "FTP Username:"
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
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1395
      End
      Begin VB.Label lblPassword 
         Caption         =   "FTP Password:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   3150
         Width           =   1395
      End
      Begin VB.Label lblHost 
         Caption         =   "FTP Host:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2340
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Current:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   3540
         Width           =   1395
      End
   End
   Begin Threed.SSPanel panButtons 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   8700
      Width           =   10485
      _Version        =   65536
      _ExtentX        =   18494
      _ExtentY        =   1085
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Tag             =   "Delete the Record Selected"
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Tag             =   "Add a new Record"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3060
         TabIndex        =   4
         Tag             =   "Update all matching records to the above"
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdEdit 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   1020
         TabIndex        =   3
         Tag             =   "Update all matching records to the above"
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Tag             =   "Update all matching records to the above"
         Top             =   120
         Width           =   870
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   750
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   6600
         Top             =   150
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
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
   End
End
Attribute VB_Name = "frmSFSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim rsSetup As New ADODB.Recordset

Private Sub LockForm(Locked As Boolean)
    cmdClose.Enabled = Locked
    cmdEdit.Enabled = Locked
    cmdOK.Enabled = Not Locked
    cmdCancel.Enabled = Not Locked
    cmdDelete.Enabled = Locked
    cmdNew.Enabled = Locked
    'clpPAYP.Enabled = Not Locked
    txtHost.Enabled = Not Locked
    txtUsername.Enabled = Not Locked
    txtPassword.Enabled = Not Locked
    chkCurrent.Enabled = Not Locked
    If Data1.Recordset.EOF Then
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()

If glbFrmCaption$ = "FTP Setup" Then
    Data1.Recordset.CancelUpdate
    'Call UpdateFormFromDB
    LockForm True
End If

If glbFrmCaption$ = "XML File Location Setup" Then 'Ticket #25604 Franks 06/18/2014
    Call XMLFileLocationScreen("DISP")
    cmdEdit.Enabled = True
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    Drive1.Enabled = False
    Dir1.Enabled = False
End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
Data1.Recordset.Delete
LockForm True
End Sub

Private Sub cmdEdit_Click()

If glbFrmCaption$ = "FTP Setup" Then
    LockForm False
End If
If glbFrmCaption$ = "XML File Location Setup" Then 'Ticket #25604 Franks 06/18/2014
    cmdEdit.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    Drive1.Enabled = True
    Dir1.Enabled = True
End If

End Sub

Private Sub cmdNew_Click()
LockForm False
Data1.Recordset.AddNew
'txtHost.Text = "sftp.payweb.ca"
txtHost.Text = "ftp01.workstreaminc.com"

End Sub

Private Sub cmdOK_Click()
Dim SQLQ As String

If glbFrmCaption$ = "FTP Setup" Then
    If Len(txtHost.Text) = 0 Then
        MsgBox "Host is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
        txtHost.SetFocus
        Exit Sub
    End If
    If Len(txtUsername.Text) = 0 Then
        MsgBox "Username is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
        txtUsername.SetFocus
        Exit Sub
    End If
    If Len(txtPassword.Text) = 0 Then
        MsgBox "Password is a required field.", vbExclamation + vbOKOnly, "Missing Required Field"
        txtPassword.SetFocus
        Exit Sub
    End If
    LockForm True
    Data1.Recordset.Update
    DoEvents
    If Not Data1.Recordset.EOF Then
        If chkCurrent.Value Then
                'update other records to non Current
                SQLQ = "UPDATE HRSF_FTP_SETUP SET CURRENT_FTP = 0 WHERE NOT [ID] = " & Data1.Recordset("ID")
                gdbAdoIhr001.Execute SQLQ
                Data1.Refresh
        Else
                SQLQ = "UPDATE HRSF_FTP_SETUP SET CURRENT_FTP = 0 WHERE  [ID] = " & Data1.Recordset("ID")
                gdbAdoIhr001.Execute SQLQ
                Data1.Refresh
        End If
    End If
End If

If glbFrmCaption$ = "XML File Location Setup" Then 'Ticket #25604 Franks 06/18/2014
    Call XMLFileLocationScreen("UPDT")
    cmdEdit.Enabled = True
    
    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    Drive1.Enabled = False
    Dir1.Enabled = False
End If

End Sub

Private Sub Dir1_Change()
    txtPDFilename.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Activate()
glbOnTop = "frmSFSetup"
Call set_Buttons
End Sub

Private Sub Form_Load()
    glbOnTop = "frmSFSetup"
    'If glbSystemData.SystemType = Access Then
    '     Data1.ConnectionString = glbSystemData.BaseConnectString & glbSystemData.Path & "IHRPWEB.MDB"
    'Else
    '     Data1.ConnectionString = glbSystemData.BaseConnectString
    'End If
    'Data1.RecordSource = "SELECT * FROM PAYWEB_SETUP"
    Me.Caption = glbFrmCaption$
    If glbFrmCaption$ = "FTP Setup" Then
        fraFTP.BorderStyle = 0
        Data1.ConnectionString = glbAdoIHRDB
        Data1.RecordSource = "SELECT * FROM HRSF_FTP_SETUP"
        Data1.Refresh
        LockForm True
        Call INI_Controls(Me)
    End If
    If glbFrmCaption$ = "XML File Location Setup" Then 'Ticket #25604 Franks 06/18/2014
        Call XMLFileLocationScreen("DISP")
    End If
End Sub
Private Sub XMLFileLocationScreen(xType)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim xPath
    
If xType = "DISP" Then
    fraFTP.Visible = False
    cmdNew.Visible = False
    cmdDelete.Visible = False
    fraXMLLocation.BorderStyle = 0
    fraXMLLocation.Top = fraFTP.Top
    fraXMLLocation.Visible = True
    SQLQ = "SELECT * FROM HRSF_XMLFILE_LOCATION"
    rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        rs.AddNew
        xPath = glbIHRREPORTS
        If Right(xPath, 1) = "\" Then
            xPath = Left(xPath, Len(xPath) - 1)
        End If
        rs("SF_LOCATION") = xPath
        rs.Update
    End If
    txtPDFilename.Text = rs("SF_LOCATION")
    If Len(Dir$(txtPDFilename.Text, vbDirectory)) > 0 Then
        Drive1.Drive = Left(txtPDFilename.Text, 2)
        Dir1.Path = txtPDFilename.Text
    End If
    rs.Close
    Drive1.Enabled = False
    Dir1.Enabled = False
End If
If xType = "UPDT" Then
    If Len(txtPDFilename.Text) > 0 Then
        SQLQ = "SELECT * FROM HRSF_XMLFILE_LOCATION"
        rs.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rs.EOF Then
            rs.AddNew
        End If
        xPath = txtPDFilename.Text
        If Right(xPath, 1) = "\" Then
            xPath = Left(xPath, Len(xPath) - 1)
        End If
        rs("SF_LOCATION") = xPath
        rs.Update
        rs.Close
    End If
End If

End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP 'MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False
End Property

Public Property Get Addable() As Boolean
Addable = False ' True
End Property
Public Property Get Updateble() As Boolean
Updateble = False '  'True
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False ' 'True
End Property

Public Property Get Printable() As Boolean
Printable = False 'True
End Property
