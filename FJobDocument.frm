VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmJobDocument 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Job Document"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   1005
   ClientWidth     =   10515
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   10515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JD_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   8760
      MaxLength       =   25
      TabIndex        =   21
      Text            =   "LUser"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JD_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   7080
      MaxLength       =   25
      TabIndex        =   20
      Text            =   "LTime"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "JD_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   5400
      MaxLength       =   25
      TabIndex        =   19
      Text            =   "Ldate"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1470
   End
   Begin VB.CommandButton cmdOpen 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   9360
      TabIndex        =   2
      Tag             =   "Close and exit this screen"
      Top             =   930
      Width           =   705
   End
   Begin VB.CommandButton cmdAttachFile 
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
      Height          =   270
      Left            =   8880
      TabIndex        =   1
      Top             =   930
      Width           =   375
   End
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "JD_COMMENT"
      Height          =   855
      Left            =   1950
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Tag             =   "00-Enter Comments"
      Top             =   1290
      Width           =   8115
   End
   Begin VB.Frame frmFile 
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   60
      TabIndex        =   11
      Top             =   60
      Width           =   10455
      Begin VB.Label lblEmpName 
         AutoSize        =   -1  'True
         Caption         =   "lblEmpName"
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
         Left            =   2640
         TabIndex        =   13
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Code/Description:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   2295
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   14
      Top             =   4755
      Width           =   10515
      _Version        =   65536
      _ExtentX        =   18547
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
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
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
         Left            =   1080
         TabIndex        =   6
         Tag             =   "Edit the Information"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
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
         Left            =   1995
         TabIndex        =   7
         Tag             =   "Save the changes made"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
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
         Left            =   2925
         TabIndex        =   8
         Tag             =   "Cancel the changes made"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "New"
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
         Left            =   3870
         TabIndex        =   9
         Tag             =   "Attach a new document"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "Delete"
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
         Left            =   4800
         TabIndex        =   10
         Tag             =   "Delete the listed document"
         Top             =   120
         Width           =   915
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "Close"
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
         TabIndex        =   5
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   915
      End
   End
   Begin TrueOleDBGrid60.TDBGrid TDBGrid1 
      Bindings        =   "FJobDocument.frx":0000
      Height          =   2265
      Left            =   120
      OleObjectBlob   =   "FJobDocument.frx":0014
      TabIndex        =   4
      Tag             =   "Province Listings"
      Top             =   2250
      Width           =   10215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "JD_DOC_TYPE"
      Height          =   285
      Index           =   0
      Left            =   1635
      TabIndex        =   0
      Tag             =   "01-Document Type Code "
      Top             =   570
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JDTC"
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      DataField       =   "JD_FILE_LINK"
      Height          =   285
      Left            =   1950
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "00-File Name (Do not Enter Extension TXT)"
      Top             =   930
      Width           =   6855
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   3600
      Top             =   4560
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin MSComDlg.CommonDialog AttachmentDialog 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblPOSID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblPOSID"
      DataField       =   "JD_JOB"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   22
      Top             =   600
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File name"
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
      Left            =   180
      TabIndex        =   18
      Top             =   975
      Width           =   780
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Document Type"
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
      Left            =   180
      TabIndex        =   16
      Top             =   600
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   180
      TabIndex        =   15
      Top             =   1290
      Width           =   870
   End
End
Attribute VB_Name = "frmJobDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim rsDATA As New ADODB.Recordset

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdAttachFile_Click()
    glbDocName = "JobFiles"
    'frmAttachFilename.Show 1
    'DoEvents
    
AttachmentDialog.DialogTitle = "Select the file to attach..."
AttachmentDialog.Filter = "*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pub;*.pdf;*.jpg|*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pub;*.pdf;*.jpg"    '"Word Documents (*.doc;*.docx)|*.doc;*.docx"
'*.doc;*.xls;*.ppt;*.pdf;*.jpg;*.docx
AttachmentDialog.FilterIndex = 1
AttachmentDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
AttachmentDialog.ShowOpen
If Len(AttachmentDialog.FileName) <> 0 Then
    txtFileName.Text = AttachmentDialog.FileName
Else
    glbDocName = ""
End If
'Remove the validation to check the file name should only consists of certain chars.
End Sub

Private Sub cmdCancel_Click()
On Error GoTo Can_Err
   
    Call Display_Value

    Call modSTUPD(False)    ' reset screen's attributes
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        cmdModify.Enabled = False
        cmdNew.Enabled = True
        cmdDelete.Enabled = False
        cmdOpen.Enabled = False
        cmdAttachFile.Enabled = False
    Else
        If Dir(txtFileName.Text) = "" Then
            cmdOpen.Enabled = False
        Else
            cmdOpen.Enabled = True
        End If
    End If
    
    fglbNew = False
    
Exit Sub
    
Can_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdCancel", "HR_JOB_DOCUMENT", "Cancel")
    Resume Next
End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdClose_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdDelete_Click()
    Dim SQLQ As String, x
    Dim Title$, Msg$, DgDef As Variant, Response%
    Dim xDocTypeDesc As String
    
    On Error GoTo Mod_Err
    
    Title = "Job File Delete"
    xDocTypeDesc = GetTABLDesc("JDTC", clpCode(0))
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Msg$ = "Are you sure you want to Delete " & lblEmpName & " Job's Document Type '" & xDocTypeDesc & "'?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        Call Display_Value
    End If
    
    fglbNew = False
        
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        cmdModify.Enabled = False
        cmdNew.Enabled = True
        cmdDelete.Enabled = False
    End If
    
    If txtFileName.Text = "" Then
        cmdOpen.Enabled = False
    Else
        cmdOpen.Enabled = True
    End If
        
    Screen.MousePointer = DEFAULT

Exit Sub

Mod_Err:
    If Err = 53 Then Resume Next
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "HR_JOB_DOCUMENT", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Sub

Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call modSTUPD(True)
clpCode(0).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdModify", "HR_JOB_DOCUMENT", "Modify")
Call RollBack  '08June99 js

End Sub

Private Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdNew_Click()
On Error GoTo NewErr

    Call modSTUPD(True)
       
    fglbNew = True
    Call Set_Control("B", Me)
           
    lblPOSID.Caption = glbPos
    cmdAttachFile.Enabled = True
    clpCode(0).SetFocus
    
    If txtFileName.Text = "" Then
        cmdOpen.Enabled = False
    Else
        cmdOpen.Enabled = True
    End If
    
Exit Sub
    
NewErr:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_JOB_DOCUMENT", "AddNew")
    Resume Next
End Sub

Private Sub cmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
    Dim intJobDocID As Integer
    On Error GoTo OK_Err
    
    If Not chkDoc() Then Exit Sub
    If fglbNew Then rsDATA.AddNew: fglbNew = False

    Call UpdUStats(Me)
    Call Set_Control("U", Me, rsDATA)
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    intJobDocID = rsDATA("JD_ID")
    gdbAdoIhr001.CommitTrans
    
    If Not glbSQL And Not glbOracle Then Call Pause(0.5)
    Data1.Refresh
    
    Data1.Recordset.Find "JD_ID=" & intJobDocID
    Call Display_Value

    fglbNew = False
       
    Call modSTUPD(False)
    
Exit Sub
    
OK_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdOK", "HR_JOB_DOCUMENT", "Update")
    Resume Next
    Unload Me
End Sub

Private Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOpen_Click()
    If Dir(txtFileName.Text) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & txtFileName.Text & "]"
        txtFileName.SetFocus
        Exit Sub
    Else
        'Open the attachment
        Shell "cmd /c " & GetShortName(txtFileName)
    End If
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    glbFrmCaption$ = Me.Caption
    glbErrNum& = ErrorNumber
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_JOB_DOCUMENT", "SELECT")
End Sub

Private Sub Form_Activate()
    'Call INI_Controls(Me)
'    Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    Dim rsEmp As New ADODB.Recordset
    Dim x%, SQLQ
    Dim Y%
    
    glbOnTop = "FRMJOBDOCUMENT"
    Screen.MousePointer = HOURGLASS
    
    Data1.ConnectionString = glbAdoIHRDB
    
    If EERetrieve() = False Then
        Screen.MousePointer = DEFAULT
        MsgBox "Error retrieving Job Files"
        Exit Sub
    End If
    
    lblPOSID = glbPos
    
    If glbDocName = "JobFiles" Or glbDocName = "EmpPosJobFiles" Then
        lblEENum(0).Caption = "Position Code/Description:"
        SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB WHERE JB_CODE='" & glbPos & "' "
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmp.EOF Then
            lblEmpName.Caption = glbPos & "/" & rsEmp("JB_DESCR")
        End If
        rsEmp.Close
    End If
        
    Call Display_Value
    Call INI_Controls(Me)
    
    Call modSTUPD(False)
    
    
    If Dir(txtFileName.Text) = "" Then
        cmdOpen.Enabled = False
    Else
        cmdOpen.Enabled = True
    End If
    
    If Data1.Recordset.EOF And Data1.Recordset.BOF Then
        cmdModify.Enabled = False
        cmdNew.Enabled = True
        cmdDelete.Enabled = False
        cmdOpen.Enabled = False
        cmdAttachFile.Enabled = False
    End If
           
    If (Not gSec_Upd_Job_Files_Attachment) Or glbDocName = "EmpPosJobFiles" Then
        cmdModify.Enabled = False
        cmdNew.Enabled = False
        cmdDelete.Enabled = False
        cmdOK.Enabled = False
        cmdCancel.Enabled = False
        cmdOpen.Enabled = False
        cmdAttachFile.Enabled = False
    End If
        
    Screen.MousePointer = DEFAULT
    
End Sub

Private Sub Form_LostFocus()
    MDIMain.panHelp(0).Caption = " "
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
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
    
    cmdOK.Enabled = TF          'May99 js
    cmdCancel.Enabled = TF      '
    cmdNew.Enabled = FT         '
    cmdClose.Enabled = FT       '
    cmdModify.Enabled = FT      '
    cmdDelete.Enabled = FT
    cmdOpen.Enabled = TF
    cmdAttachFile.Enabled = TF
    TDBGrid1.Enabled = FT

    clpCode(0).Enabled = TF
    memComments.Enabled = TF
    txtFileName.Enabled = TF
    
    If Dir(txtFileName.Text) = "" Then
        cmdOpen.Enabled = False
    Else
        cmdOpen.Enabled = True
    End If
End Sub

Private Sub TDBGrid1_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    If TDBGrid1.Tag = "ASC" Then
        TDBGrid1.Tag = "DESC"
    Else
        TDBGrid1.Tag = "ASC"
    End If
    
    SQLQ = "SELECT * FROM HR_JOB_DOCUMENT "
    SQLQ = SQLQ & " WHERE JD_JOB ='" & glbPos & "' "
    SQLQ = SQLQ & " ORDER BY " & TDBGrid1.Columns(ColIndex).DataField & " " & TDBGrid1.Tag
    
    Data1.RecordSource = SQLQ
    Data1.Refresh
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Sub txtFileName_Change()
    If Dir(txtFileName.Text) = "" Then
        cmdOpen.Enabled = False
    Else
        cmdOpen.Enabled = True
    End If
End Sub

Sub txtFileName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Sub txtFileName_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Function chkDoc()
    Dim Alphabet, xlen, i%, xwk, xok
    Dim strActFilename As String
    On Error GoTo chkDoc_Err
    
    chkDoc = False
        
    If Not clpCode(0).ListChecker Then Exit Function
    
    If Len(txtFileName) = 0 Then
        MsgBox "File Name is required."
        txtFileName.SetFocus
        Exit Function
    End If
        
    If Dir(txtFileName.Text) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & txtFileName.Text & "]"
        txtFileName.SetFocus
        Exit Function
    End If

    chkDoc = True
    
Exit Function
    
chkDoc_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkDoc", "HR_JOB_DOCUMENT", "Edit/Add")
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If
End Function

Public Property Get ChangeAction() As UpdateStateEnum
    ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
    RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
    UpdateRight = True
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
    Printable = False
End Property

Public Sub SET_UP_MODE()
    Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
    Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Public Function EERetrieve()
    Dim SQLQ$
    
    EERetrieve = False
    Screen.MousePointer = HOURGLASS
    
    On Error GoTo EERetrieveErr
    
    SQLQ$ = "SELECT * FROM HR_JOB_DOCUMENT "
    SQLQ$ = SQLQ$ & "WHERE JD_JOB = '" & glbPos & "'"
    SQLQ$ = SQLQ$ & "ORDER BY JD_DOC_TYPE"
    
    Data1.RecordSource = SQLQ$
    Data1.Refresh
    
    
    EERetrieve = True
    Screen.MousePointer = DEFAULT
    
    Exit Function
    
EERetrieveErr:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_JOB_DOCUMENT", "SELECT")
    Call RollBack

End Function

Public Sub Display_Value()
Dim SQLQ

If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
    Call Set_Control("B", Me)
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "SELECT * FROM HR_JOB_DOCUMENT "
    SQLQ = SQLQ & " WHERE JD_ID = " & Data1.Recordset!JD_ID
    If rsDATA.State <> 0 Then rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End If

End Sub
