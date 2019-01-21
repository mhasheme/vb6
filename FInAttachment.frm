VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmInAttachment 
   AutoRedraw      =   -1  'True
   Caption         =   "Import Attachment"
   ClientHeight    =   3180
   ClientLeft      =   15
   ClientTop       =   1020
   ClientWidth     =   8760
   ForeColor       =   &H00000000&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog AttachmentDialog 
      Left            =   360
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2100
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4860
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2100
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   5040
      MultiSelect     =   2  'Extended
      Pattern         =   "*.doc;*.xls;*.ppt;*.pdf;*.jpg;*.docx"
      TabIndex        =   10
      Top             =   4530
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Frame frmFile 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   60
      TabIndex        =   12
      Top             =   60
      Width           =   8655
      Begin VB.TextBox txtDocDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   3
         Tag             =   "00-Document Description"
         Top             =   1440
         Width           =   5895
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
         Left            =   7980
         TabIndex        =   1
         Tag             =   "Click to select the location"
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   0
         Tag             =   "00-File Name (Do not Enter Extension TXT)"
         Top             =   720
         Width           =   5895
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   0
         Left            =   1720
         TabIndex        =   2
         Tag             =   "01-Document Type Code "
         Top             =   1080
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "DOCT"
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee # / Name:"
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
         TabIndex        =   14
         Top             =   120
         Width           =   1755
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1125
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Description"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Label lblDisp 
         Alignment       =   2  'Center
         Caption         =   "Label"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   8175
      End
      Begin VB.Label lblEmpName 
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
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   120
         Width           =   5175
      End
      Begin VB.Label lblEmpNo 
         AutoSize        =   -1  'True
         Caption         =   "lblEmpNo"
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
         Left            =   2040
         TabIndex        =   15
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Import File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   765
         Width           =   1185
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   2520
      Width           =   8760
      _Version        =   65536
      _ExtentX        =   15452
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
      Begin VB.CommandButton cmdSave 
         Appearance      =   0  'Flat
         Caption         =   "Save Type && Description"
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
         Left            =   3840
         TabIndex        =   7
         Tag             =   "Save Document Type and Description"
         Top             =   150
         Width           =   2295
      End
      Begin VB.CommandButton cmdModify 
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
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Tag             =   "Import New Attachment"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
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
         Left            =   960
         TabIndex        =   5
         Tag             =   "Delete Attachment"
         Top             =   150
         Width           =   825
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
         Left            =   1920
         TabIndex        =   6
         Tag             =   "Close Import Attachment"
         Top             =   150
         Width           =   735
      End
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Import From Path"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   3060
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmInAttachment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FPath ', UPDTCNT
Dim ImportFile As String
Dim xDeleteable
Dim chkFile As Integer
Dim flgWrongDocTypeCode As Boolean

Private Sub clpCode_LostFocus(Index As Integer)
    If Len(clpCode(0).Text) > 0 Then
        Dim rs As New ADODB.Recordset
        Dim strSQL As String
        Dim xWrongPos, xPos, I
        Dim xList, xShowCell, xCell
        Dim xTemplate As String
        
        If clpCode(0).Caption = "Unassigned" Then Exit Sub
        
        '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
        xTemplate = ""
        xTemplate = Get_Template(glbUserID)
        
        
        xList = clpCode(0).Text
        xWrongPos = 0
        xPos = 0
        Do While Len(xList) <> 0
            xWrongPos = xWrongPos + xPos
            xPos = InStr(xList, ",")
            If xPos = 0 Then
                xShowCell = xList
                xList = ""
            Else
                xShowCell = Left(xList, xPos - 1)
                xList = Mid(xList, xPos + 1)
            End If
            xCell = xShowCell
            
            If xTemplate = "" Or xTemplate = "TEMPLATE" Then
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
            Else
                '????Ticket #24808 -  Retrieve template's security profile
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            End If
            strSQL = strSQL & " AND CODENAME = '" & xCell & "' AND TB_NAME='DOCT'"
            rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False And rs.BOF = False Then
                If rs("ACCESSABLE") = 0 Then
                    flgWrongDocTypeCode = True
                    MsgBox "You do not have Authorization for '" & xCell & "' Document Type code", vbInformation + vbOKOnly, "Authorization Failure"
                    SendKeys "{HOME}"
                    For I = 1 To xWrongPos
                        SendKeys "{Right}"
                    Next
                    Exit Sub
                End If
            Else
                flgWrongDocTypeCode = True
                MsgBox "You do not have Authorization for '" & xCell & "' Document Type code", vbInformation + vbOKOnly, "Authorization Failure"
                SendKeys "{HOME}"
                For I = 1 To xWrongPos
                    SendKeys "{Right}"
                Next
                Exit Sub
            End If
            rs.Close
            Set rs = Nothing
        Loop
    End If

End Sub

Private Sub cmdBrowse_Click()
AttachmentDialog.DialogTitle = "Select the file to attach..."
If glbDocName = "INJURYWF7" Or glbDocName = "INJURYWF7_WRITTENOFR" Then
    AttachmentDialog.Filter = "*.pdf|*.pdf"    '"Word Documents (*.doc;*.docx)|*.doc;*.docx"
Else
    AttachmentDialog.Filter = "*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pub;*.pdf;*.jpg|*.doc;*.docx;*.xls;*.xlsx;*.ppt;*.pptx;*.pub;*.pdf;*.jpg"    '"Word Documents (*.doc;*.docx)|*.doc;*.docx"
End If
AttachmentDialog.FilterIndex = 1
AttachmentDialog.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
AttachmentDialog.ShowOpen
If Len(AttachmentDialog.FileName) <> 0 Then
    txtFileName.Text = AttachmentDialog.FileName
End If
'Remove the validation to check the file name should only consists of certain chars.
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Private Sub cmdDelete_Click()
Dim SQLQ As String, X
Dim Title$, Msg$, DgDef As Variant, Response%

Title = "Employee " & glbDocName & " Delete"

On Error GoTo Mod_Err
    
'Release 8.1
'Check if the User have Security rights on this Document Type Code
'Send LostFocus on Document Type code so it is validated as per the Document Type Codes security
flgWrongDocTypeCode = False
Call clpCode_LostFocus(0)
If flgWrongDocTypeCode = True Then Exit Sub

    
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are you sure you want to Delete " & lblEmpName & "'s " & glbDocName & "?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then Exit Sub

Screen.MousePointer = HOURGLASS
Select Case glbDocName
    Case "Resume"
        'gdbAdoIhr001.Execute "delete from HRDOC_EMP WHERE RE_TYPE='RESUME' AND RE_EMPNBR=" & glbLEE_ID
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_EMP WHERE RE_TYPE='RESUME' AND RE_EMPNBR=" & glbLEE_ID
    
    Case "Offer"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR=" & glbLEE_ID & " AND DJ_JOB= '" & glbJob & "' AND DJ_SDATE =" & Date_SQL(glbSDate)
    
    Case "Jobdescription"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_JOB WHERE DB_TYPE='" & UCase(glbDocName) & "' AND DB_JOB= '" & glbPos & "'"
    
    Case "Counsel"
        If glbtermopen Then
            gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DC_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE Term_HR_COUNSEL SET CL_DOCKEY = Null WHERE CL_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND CL_DOCKEY= " & glbDocKey & " "
        Else
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR=" & glbLEE_ID & " AND DC_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HR_COUNSEL SET CL_DOCKEY = Null WHERE CL_EMPNBR=" & glbLEE_ID & " AND CL_DOCKEY= " & glbDocKey & " "
        End If
    Case "Comments"
        If glbtermopen Then
            gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND DO_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE Term_COMMENTS SET CO_DOCKEY = Null WHERE CO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND CO_DOCKEY= " & glbDocKey & " "
        Else
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbLEE_ID & " AND DO_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HR_COMMENTS SET CO_DOCKEY = Null WHERE CO_EMPNBR=" & glbLEE_ID & " AND CO_DOCKEY= " & glbDocKey & " "
        End If
    Case "INCIDENT"
        'gdbAdoIhr001_DOC.Execute "delete from HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbLEE_ID & " AND DO_DOCKEY= " & glbDocKey & " "
        'gdbAdoIhr001_DOC.Execute "Update HRDOC_HEALTH_SAFETY set DE_FILEEXT = null WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID & " AND DE_CASE= '" & glbJob & "' AND DE_DOCNO ='" & frmEHSAttach.txtDocNum & "'"
        SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND DE_CASE= '" & glbJob & "'"
        SQLQ = SQLQ & " AND DE_DOCNO= '" & glbDocTmp & "'"
        gdbAdoIhr001_DOC.Execute SQLQ
    
    Case "INJURYWF7"
        'gdbAdoIhr001_DOC.Execute "delete from HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbLEE_ID & " AND DO_DOCKEY= " & glbDocKey & " "
        'gdbAdoIhr001_DOC.Execute "Update HRDOC_HEALTH_SAFETY set DE_FILEEXT = null WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID & " AND DE_CASE= '" & glbJob & "' AND DE_DOCNO ='" & frmEHSAttach.txtDocNum & "'"
        SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='" & UCase(glbDocName) & "' AND W7_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND W7_CASE = '" & glbJob & "'"
        SQLQ = SQLQ & " AND W7_DOCKEY = '" & glbDocKey & "'"
        gdbAdoIhr001_DOC.Execute SQLQ
        
        'Ticket #25355 - Remove the link to the master table
        SQLQ = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_DOCKEY = Null WHERE EC_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND EC_CASE = '" & glbJob & "'"
        SQLQ = SQLQ & " AND EC_DOCKEY = '" & glbDocKey & "'"
        gdbAdoIhr001.Execute SQLQ
        
    Case "INJURYWF7_WRITTENOFR"
        SQLQ = "DELETE FROM HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='" & UCase(glbDocName) & "' AND F7_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND F7_CASE = '" & glbJob & "'"
        SQLQ = SQLQ & " AND F7_DOCKEY = '" & glbDocKey & "'"
        gdbAdoIhr001_DOC.Execute SQLQ
        
        'Ticket #25355 - Remove the link to the master table
        SQLQ = "UPDATE HR_OHS_FORM7_SECTIONS SET F7_DOCKEY = Null WHERE F7_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND F7_CASE = '" & glbJob & "'"
        SQLQ = SQLQ & " AND F7_DOCKEY = '" & glbDocKey & "'"
        gdbAdoIhr001.Execute SQLQ
        
    Case "Performance"
        'gdbAdoIhr001_DOC.Execute "delete from HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(glbDocName) & "' AND DH_EMPNBR=" & glbLEE_ID & " AND DH_JOB= '" & glbJob & "' AND DH_PREVDATE =" & Date_SQL(glbSDate)
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(glbDocName) & "' AND DH_EMPNBR=" & glbLEE_ID & " AND DH_DOCKEY= " & glbDocKey & " "
        
        'Ticket #25355 - Remove the link to the master table
        gdbAdoIhr001.Execute "UPDATE HR_PERFORM_HISTORY SET PH_DOCKEY = Null WHERE PH_EMPNBR=" & glbLEE_ID & " AND PH_DOCKEY= " & glbDocKey & " "
        
    Case "EdSem"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_EDSEM WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID & " AND ES_DOCKEY= " & glbDocKey & " "
        
        'Ticket #25355 - Remove the link to the master table
        gdbAdoIhr001.Execute "UPDATE HREDSEM SET ES_DOCKEY = Null WHERE ES_EMPNBR=" & glbLEE_ID & " AND ES_DOCKEY= " & glbDocKey & " "
        
    Case "EdSem_Retest"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_EDSEM_RETEST WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID & " AND ES_DOCKEY= " & glbDocKey & " "
        
        'Ticket #25355 - Remove the link to the master table
        gdbAdoIhr001.Execute "UPDATE HREDSEM_RETEST SET ES_DOCKEY = Null WHERE ES_EMPNBR=" & glbLEE_ID & " AND ES_DOCKEY= " & glbDocKey & " "
        
    Case "FormalEdu"
        If glbtermopen Then
            gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE Term_EDU SET EU_DOCKEY = Null WHERE EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY= " & glbDocKey & " "
        Else
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR=" & glbLEE_ID & " AND EU_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HREDU SET EU_DOCKEY = Null WHERE EU_EMPNBR=" & glbLEE_ID & " AND EU_DOCKEY= " & glbDocKey & " "
        End If
    Case "DollarEnt"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_HRDOLENT WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID & " AND DE_DOCKEY= " & glbDocKey & " "
        
        'Ticket #25355 - Remove the link to the master table
        gdbAdoIhr001.Execute "UPDATE HRDOLENT SET DE_DOCKEY = Null WHERE DE_EMPNBR=" & glbLEE_ID & " AND DE_DOCKEY= " & glbDocKey & " "
        
    Case "Associations"
        If glbtermopen Then
            gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND TD_CODE='" & glbAssocCode & "' AND TD_BEGINDT=" & Date_SQL(glbBeginDt)    '" AND TD_DOCKEY= " & glbDocKey & " "
        Else
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR=" & glbLEE_ID & " AND TD_CODE='" & glbAssocCode & "' AND TD_BEGINDT=" & Date_SQL(glbBeginDt)    '" AND TD_DOCKEY= " & glbDocKey & " "
        End If
    Case "Attendance"
        If glbtermopen Then
            gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & glbDocKey & " "
                    
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE Term_ATTENDANCE SET AD_DOCKEY = Null WHERE AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & glbDocKey & " "
            
        Else
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR=" & glbLEE_ID & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HR_ATTENDANCE SET AD_DOCKEY = Null WHERE AD_EMPNBR=" & glbLEE_ID & " AND AD_REASON='" & glbAttReason & "' AND AD_DOA=" & Date_SQL(glbAttDOA) & " AND AD_DOCKEY= " & glbDocKey & " "
            
        End If
    Case "Termination"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_EMP WHERE RE_TYPE='TERMINATION' AND RE_EMPNBR=" & glbLEE_ID
    
    Case "EmployeeFlag"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='EMPLOYEEFLAG' AND EF_EMPNBR=" & glbLEE_ID
    
    Case "PositionSkill"
        gdbAdoIhr001_DOC.Execute "delete from HRDOC_JOBSKL WHERE DS_TYPE='" & UCase(glbDocName) & "' AND DS_JOB= '" & glbPos & "' AND DS_SKILL= '" & glbPosSkill & "'"

    'Release 8.1
    Case "OtherInfo"
        If glbtermopen Then
            gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
        Else
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbLEE_ID
        End If

    'Release 8.1
    Case "LOA"
        'If glbtermopen Then
        '    gdbAdoIhr001_DOC.Execute "delete from Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
        'Else
            gdbAdoIhr001_DOC.Execute "delete from HRDOC_HRSTATUS WHERE SC_TYPE='" & UCase(glbDocName) & "' AND SC_EMPNBR=" & glbLEE_ID & " AND SC_DOCKEY= " & glbDocKey & " "
            
            'Ticket #25355 - Remove the link to the master table
            gdbAdoIhr001.Execute "UPDATE HRSTATUS SET SC_DOCKEY = Null WHERE SC_EMPNBR=" & glbLEE_ID & " AND SC_DOCKEY= " & glbDocKey & " "
        'End If

End Select

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Unload Me

Exit Sub

Mod_Err:
If Err = 53 Then Resume Next

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "Delete Attachment", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Sub cmdModify_Click()
Dim SQLQ As String, X
Dim Title$, Msg$, DgDef As Variant, Response%
Dim xDocName As String

'George Jan 19,2006 begin
'    If glbDocName = "Resume" Then
'        Title = "Employee Resume Import"
'    End If

    Select Case glbDocName 'George Jan 19,2006
        Case "Resume"
            Title = "Employee Resume Import"
        Case "Offer"
            Title = "Job Offer Import"
        Case "OfferMU"
            'Ticket #28457 City of Niagara Falls
            If glbCompSerial = "S/N - 2276W" Then
                Title = "Lay Off / Return from Lay Off Letter Import"
            Else
                Title = "Job Offer Import"
            End If
        Case "Jobdescription"
            Title = "Job Description Import"
        Case "Comments"
            Title = "Comments Import"
        Case "INCIDENT"
            Title = "Incidents Import"
        Case "INJURYWF7"
            Title = "WSIB Form 7 - Concerns about the Claim Import"
        Case "INJURYWF7_WRITTENOFR"
            Title = "WSIB Form 7 - Written Offer given to the worker Import"
        Case "Counsel"
            Title = "Counsel Import"
        Case "Performance"
            Title = "Performance Review Import"
        Case "EdSem"
            Title = "Continuing Education Import"
        Case "EdSem_Retest"
            Title = "Continuing Education Retest Import"
        Case "FormalEdu"
            Title = "Formal Education Import"
        Case "DollarEnt"
            Title = "Dollar Entitlement Import"
        Case "Associations"
            Title = "Associations Import"
        Case "Attendance"
            Title = "Attendance Import"
        Case "AttendanceMU"
            Title = "Attendance Import"
        Case "Termination"
            Title = "Termination Import"
        Case "EmployeeFlag"
            Title = "Employee Flag Import - " & glbEmpFlag
        Case "PositionSkill"
            Title = "Position Skill Import"
        'Release 8.1
        Case "OtherInfo"
            Title = "Other Information Import"
        Case "LOA"
            Title = "Leave of Absence Import"
    
    End Select
'George Jan 19,2006 end
    
    On Error GoTo Mod_Err
    
    If Not chkDoc() Then Exit Sub
    
    '8.0 - Ticket #22682 - If Document Description missing put Document Type + Date
    If Len(Trim(txtDocDesc.Text)) = 0 Then
        txtDocDesc.Text = clpCode(0).Text & " - " & Format(Now, "mm/dd/yyyy")
    End If
    
    xDocName = glbDocName
    If glbDocName = "EdSem" Then
        xDocName = "Continuing Education"
    End If
    If glbDocName = "EdSem_Retest" Then
        xDocName = "Continuing Education Retest"
    End If
    If glbDocName = "FormalEdu" Then
        xDocName = "Formal Education"
    End If
    If glbDocName = "DollarEnt" Then
        xDocName = "Dollar Entitlement"
    End If
    If glbDocName = "INJURYWF7" Then
        xDocName = "Concerns about the claim written submission"
    End If
    If glbDocName = "INJURYWF7_WRITTENOFR" Then
        xDocName = "Written Offer given to the worker"
    End If
    If glbDocName = "PositionSkill" Then
        xDocName = "Position Skill document"
    End If
    'Release 8.1
    If glbDocName = "OtherInfo" Then
        xDocName = "Other Information"
    End If
    If glbDocName = "LOA" Then
        xDocName = "LOA Information"
    End If
    If glbDocName = "OfferMU" Then
        'Ticket #28457 City of Niagara Falls
        If glbCompSerial = "S/N - 2276W" Then
            xDocName = "Lay off / Return from Lay off Information"
        Else
            xDocName = "Job Information"
        End If
    End If
    If glbDocName = "AttendanceMU" Then
        xDocName = "Attendance Document"
    End If
    
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    If Len(lblDisp.Caption) > 0 Then
        Msg$ = "The attachment exists already." & Chr(10)
        Msg$ = Msg$ & "Are you sure you want to overwrite this " & xDocName & "? "
    Else
        Msg$ = Msg$ & "Are you sure you want to Import " & xDocName & "? "
    End If
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    
    ChDir FPath
    
    If glbDocName = "Resume" Or glbDocName = "Offer" Or glbDocName = "OfferMU" _
        Or glbDocName = "Performance" Or glbDocName = "Comments" _
        Or glbDocName = "Counsel" Or glbDocName = "Jobdescription" _
        Or glbDocName = "INCIDENT" Or glbDocName = "INJURYWF7" Or glbDocName = "EdSem" _
        Or glbDocName = "EdSem_Retest" Or glbDocName = "FormalEdu" Or glbDocName = "DollarEnt" _
        Or glbDocName = "Associations" Or glbDocName = "Attendance" Or glbDocName = "AttendanceMU" Or glbDocName = "Termination" _
        Or glbDocName = "EmployeeFlag" Or glbDocName = "INJURYWF7_WRITTENOFR" Or glbDocName = "PositionSkill" _
        Or glbDocName = "OtherInfo" Or glbDocName = "LOA" Then
        If Not modUpdateSelectionResume() Then GoTo bpMod
    End If
    
    MDIMain.panHelp(0).FloodPercent = 100
    
    Close
    '-----------------------------------------------------
    
    Screen.MousePointer = DEFAULT
    MDIMain.panHelp(0).FloodType = 0
    'MDIMain.panHelp(1).Caption = " Update Completed"
    MDIMain.panHelp(2).Caption = ""
    'If glbUPDTCNT = 0 Then
    '    Msg$ = "No " & glbDocName & " Imported "
    'Else
    '    'Msg$ = str(glbUPDTCNT)
    '    Msg$ = "This File has been Imported Successfully. "
    'End If
    DgDef = MB_ICONINFORMATION
    'MsgBox Msg$, DgDef, Title

bpMod:

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Unload Me

Exit Sub

Mod_Err:
If Err = 53 Then Resume Next

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Attach Document", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub cmdSave_Click()
    Dim SQLQ
    Dim rsTemp As New ADODB.Recordset
    Dim fldPrx As String
    Dim Title$, Msg$, DgDef As Variant, Response%

    'Confirm Save
    'Title = "Save Document Type and Description"
    'DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    'Msg$ = "Are you sure you want to Save " & lblEmpName & "'s " & glbDocName & " Document Type and Description?"
    'Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    'If Response% = IDNO Then Exit Sub
    
    fldPrx = ""
    
    If (Len(glbDocKey) = 0 And Len(glbJob) = 0) And (Not glbDocName = "Resume") And (Not glbDocName = "Termination") And (Not glbDocName = "OtherInfo") Then    'Release 8.1
        Exit Sub
    ElseIf (Len(glbDocKey) = 0 Or glbDocKey = 0) And (glbDocName <> "Resume") And (glbDocName <> "Termination") And (glbDocName <> "Offer") And (glbDocName <> "Jobdescription") And (glbDocName <> "INCIDENT") And (glbDocName <> "EmployeeFlag") And (glbDocName <> "PositionSkill") And (glbDocName <> "OtherInfo") Then
        Exit Sub
    End If
    
    Select Case glbDocName
        Case "Resume"
            SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(glbDocName) & "' AND RE_EMPNBR=" & glbLEE_ID
            'RsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            fldPrx = "RE"
        Case "Offer"
            SQLQ = "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND DJ_JOB= '" & glbJob & "'"
            SQLQ = SQLQ & " AND DJ_SDATE= " & Date_SQL(glbSDate)
            fldPrx = "DJ"
        Case "Jobdescription"
            SQLQ = "SELECT * FROM HRDOC_JOB WHERE DB_TYPE='" & UCase(glbDocName) & "'"
            SQLQ = SQLQ & " AND DB_JOB= '" & glbPos & "'"
            fldPrx = "DB"
        Case "Counsel"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR=" & glbLEE_ID
                'SQLQ = SQLQ & " AND DC_CLTYPE= '" & glbJob & "'"
                'SQLQ = SQLQ & " AND DC_COUDATE= " & Date_SQL(glbSDate)
            End If
            SQLQ = SQLQ & " AND DC_DOCKEY= " & glbDocKey & ""
            fldPrx = "DC"
        Case "Comments"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbLEE_ID
                'SQLQ = SQLQ & " AND DO_COTYPE= '" & glbJob & "'"
                'SQLQ = SQLQ & " AND DO_EDATE= " & Date_SQL(glbSDate)
            End If
            SQLQ = SQLQ & " AND DO_DOCKEY= " & glbDocKey & ""
            fldPrx = "DO"
        Case "INCIDENT"
            SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND DE_CASE= '" & glbJob & "'"
            SQLQ = SQLQ & " AND DE_DOCNO= '" & glbDocTmp & "'"
            fldPrx = "DE"
        Case "INJURYWF7"
            SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='" & UCase(glbDocName) & "' AND W7_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND W7_CASE= '" & glbJob & "'"
            SQLQ = SQLQ & " AND W7_DOCKEY= '" & glbDocKey & "'"
            fldPrx = "W7"
        Case "INJURYWF7_WRITTENOFR"
            SQLQ = "SELECT * FROM HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='" & UCase(glbDocName) & "' AND F7_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND F7_CASE= '" & glbJob & "'"
            SQLQ = SQLQ & " AND F7_DOCKEY= '" & glbDocKey & "'"
            fldPrx = "F7"
        Case "Performance"
            SQLQ = "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(glbDocName) & "' AND DH_EMPNBR=" & glbLEE_ID
            'SQLQ = SQLQ & " AND DH_JOB= '" & glbJob & "'"
            'SQLQ = SQLQ & " AND DH_PREVDATE= " & Date_SQL(glbSDate)
            SQLQ = SQLQ & " AND DH_DOCKEY= " & glbDocKey & ""
            fldPrx = "DH"
        Case "EdSem"
            SQLQ = "SELECT * FROM HRDOC_EDSEM WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND ES_DOCKEY= " & glbDocKey & ""
            fldPrx = "ES"
        Case "EdSem_Retest"
            SQLQ = "SELECT * FROM HRDOC_EDSEM_RETEST WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND ES_DOCKEY= " & glbDocKey & ""
            fldPrx = "ES"
        Case "FormalEdu"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR=" & glbLEE_ID
            End If
            SQLQ = SQLQ & " AND EU_DOCKEY= " & glbDocKey & ""
            fldPrx = "EU"
        Case "DollarEnt"
            SQLQ = "SELECT * FROM HRDOC_HRDOLENT WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND DE_DOCKEY= " & glbDocKey & ""
            fldPrx = "DE"
        Case "Associations"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR=" & glbLEE_ID
            End If
            SQLQ = SQLQ & " AND TD_DOCKEY= " & glbDocKey & ""
            fldPrx = "TD"
        Case "Attendance"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR=" & glbLEE_ID
            End If
            SQLQ = SQLQ & " AND AD_DOCKEY= " & glbDocKey & ""
            fldPrx = "AD"
        Case "Termination"
            SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(glbDocName) & "' AND RE_EMPNBR=" & glbLEE_ID
            fldPrx = "RE"
        Case "EmployeeFlag"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(glbDocName) & "' AND EF_EMPNBR=" & glbLEE_ID
            End If
            fldPrx = "EF"
        Case "PositionSkill"
            SQLQ = "SELECT * FROM HRDOC_JOBSKL WHERE DS_TYPE='" & UCase(glbDocName) & "'"
            SQLQ = SQLQ & " AND DS_JOB= '" & glbPos & "'"
            SQLQ = SQLQ & " AND DS_SKILL= '" & glbPosSkill & "'"
            fldPrx = "DS"
        Case "OtherInfo"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_HREMP_OTHER  WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbLEE_ID
            End If
            fldPrx = "ER"
        Case "LOA"
            'If glbtermopen Then
            '    SQLQ = "SELECT * FROM Term_HRDOC_HREMP_OTHER  WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            'Else
                SQLQ = "SELECT * FROM HRDOC_HRSTATUS WHERE SC_TYPE='" & UCase(glbDocName) & "' AND SC_EMPNBR=" & glbLEE_ID
            'End If
            SQLQ = SQLQ & " AND SC_DOCKEY= " & glbDocKey & ""
            fldPrx = "SC"
    End Select
    
    rsTemp.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp(fldPrx & "_DOCTYPE")) Then
            'Save Document Type and Description
            rsTemp(fldPrx & "_DOCTYPE") = clpCode(0).Text
            IIf(IsNull(rsTemp(fldPrx & "_USRDESC")), "", rsTemp(fldPrx & "_USRDESC")) = txtDocDesc.Text
            rsTemp.Update
            
            'Message
            MsgBox "Document Type and Description Saved.", vbOKOnly, "Saved the changes"
        End If
    End If
    rsTemp.Close
    Set rsTemp = Nothing
End Sub

Private Sub Dir1_Change()
    ChDir Dir1.Path
    File1.Path = Dir1.Path
    'File1.Pattern = "*.JPG"
End Sub

Private Sub Drive1_Change()
Dim xdir, xerror
On Error GoTo CKERROR
xerror = False
Dir1.Path = Drive1.Drive
Exit Sub
CKERROR:
    If Err = 68 Then
         MsgBox "Invalid Drive Selected"
         Drive1.Drive = App.Path
         xerror = True
         Resume Next
    End If
    MsgBox "ERROR " & Str(Err)
    xerror = True
    Resume Next

End Sub

Private Sub File1_Click()
    Dim iit As Integer
    Dim ii1 As Long
    Dim sit As String
    For iit = 0 To File1.ListCount - 1
        If File1.selected(iit) Then
            sit = File1.List(iit)
            txtFileName.Text = UCase(File1.List(iit))
        End If
    Next

End Sub

Private Sub Form_Activate()
Call INI_Controls(Me)
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim rsEmp As New ADODB.Recordset
Dim X%, SQLQ
Dim Y%

glbOnTop = "FRMINATTACHMENT"

Screen.MousePointer = HOURGLASS
Screen.MousePointer = DEFAULT

'Changed for WFC to w: for rest will be c:
If glbWFC Then
    Drive1.Drive = "w:"
    Dir1.Path = "w:\"
Else
    Drive1.Drive = "c:"
    Dir1.Path = "c:\"
End If
FPath = Dir1.Path

If glbDocName = "Jobdescription" Or glbDocName = "PositionSkill" Then
    lblEENum(0).Caption = "Position Code/Description:"
    SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB WHERE JB_CODE='" & glbPos & "' "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        lblEmpNo.Caption = "" 'glbPos
        If glbDocName = "Jobdescription" Then
            lblEmpName.Caption = glbPos & " / " & rsEmp("JB_DESCR")
        ElseIf glbDocName = "PositionSkill" Then
            lblEmpName.Caption = glbPos & " / " & rsEmp("JB_DESCR") & " - " & glbPosSkill & "/" & GetTABLDesc("EDSK", glbPosSkill)
        End If
    End If
    rsEmp.Close
Else
    If Not glbtermopen Then
        SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Else
        SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME FROM TERM_HREMP WHERE ED_EMPNBR=" & glbTERM_ID
        SQLQ = SQLQ & " AND TERM_SEQ = " & glbTERM_Seq
        rsEmp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    End If
    'rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not glbtermopen Then
            lblEmpNo.Caption = ShowEmpnbr(glbLEE_ID)
        Else
            lblEmpNo.Caption = ShowEmpnbr(glbTERM_ID)
        End If
        lblEmpName.Caption = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
        
        'Ticket #28457 City of Niagara Falls
        If (glbDocName = "OfferMU" Or glbDocName = "AttendanceMU") Then
            lblEmpNo.Caption = "Mass Update"
            lblEmpName.Caption = ""
        End If
    Else
        lblEmpNo.Caption = ""
        lblEmpName.Caption = ""
    
        'Ticket #28457 City of Niagara Falls
        If (glbDocName = "OfferMU" Or glbDocName = "AttendanceMU") Then
            lblEmpNo.Caption = "Mass Update"
            lblEmpName.Caption = ""
        End If
    End If
    rsEmp.Close
End If

If IfExist Then
    lblDisp.Caption = "The attachment already exists."
    
    'Ticket #28839 - Allow Document Type and Description to be changed and saved
    cmdSave.Enabled = True
Else
    lblDisp.Caption = ""
    cmdDelete.Enabled = False
    
    'Ticket #28839 - Allow Document Type and Description to be changed and saved
    cmdSave.Enabled = False
        
    'Ticket #28457 City of Niagara Falls
    If (glbDocName = "OfferMU" Or glbDocName = "AttendanceMU") And glbDocImpFile <> "" Then
        txtFileName.Text = glbDocImpFile
        clpCode(0).Text = glbDocType
        txtDocDesc.Text = glbDocDesc
    End If
End If

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

Sub txtFileName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub txtFileName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Function modUpdateSelectionResume()
Dim xxx, xx1, X%, XCNT
Dim xEmpnbr, xShowEmpNbr
Dim SQLQ
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%, SSERIAL
Dim rsEmp As New ADODB.Recordset
Dim xPath, xFileName As String

On Error GoTo modUpdateSelection_Err

modUpdateSelectionResume = False

glbUPDTCNT = 0
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(0).FloodType = 1

MDIMain.panHelp(0).FloodPercent = 0


'xFileName = UCase(txtFileName)
xFileName = ImportFile
glbDocImpFile = ImportFile

If glbtermopen Then
    xEmpnbr = glbTERM_ID
Else
    xEmpnbr = glbLEE_ID
End If
If Not IsNumeric(xEmpnbr) Then xEmpnbr = 0

'Not new record, save the file into database here
If Not glbDocNewRecord Then
    Call AttachmentAdd(xEmpnbr, xFileName, clpCode(0).Text, txtDocDesc.Text)
End If

'8.0 - Ticket #22682 - For new records you need these values
glbDocType = clpCode(0).Text
glbDocDesc = txtDocDesc.Text

MDIMain.panHelp(0).Caption = ""
modUpdateSelectionResume = True
Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:
    Screen.MousePointer = DEFAULT
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update", "Add Attachment", "Document")
    Screen.MousePointer = DEFAULT
    
    If gintRollBack% = False Then Resume Next Else Unload Me
End Function

Function chkDoc()
Dim Alphabet, xlen, I%, xwk, xok
chkDoc = False
On Error GoTo chkDoc_Err


    If Len(txtFileName) = 0 Then
        MsgBox "File Name is required."
        txtFileName.SetFocus
        Exit Function
    End If
    
    txtFileName = LTrim(txtFileName)
    xlen = Len(txtFileName)
    ' dkostka - 10/16/2001 - Added space and -_()! to end of alphabet, filenames can have these chars
    'Hemu - Ticket #16031 - With French accents - àâäæçéèêëîïôœùûü«€ÀÂÄÆÇÉÈÊËÎÏÔŒÙÛÜ»
'    Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-_()!., àâäæçéèêëîïôœùûü«€ÀÂÄÆÇÉÈÊËÎÏÔŒÙÛÜ»"

'    xok = True
'    For i% = 1 To xlen
'        xwk = Mid(txtFileName, i%, 1)
'        If InStr(Alphabet, xwk) = 0 Then
'            xok = False
'            Exit For
'        End If
'    Next
'    If Not xok Then
'        MsgBox "Invalid File Name"
'        txtFileName.SetFocus
'        Exit Function
'    End If

'    ImportFile = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\")) & txtFileName
    ImportFile = txtFileName
    'MsgBox ImportFile
    If Dir(ImportFile) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]"
        txtFileName.SetFocus
        Exit Function
    End If

    '8.0 - Ticket #22682 - Add Document Type table code
    If Not clpCode(0).ListChecker Then Exit Function
        
    If Len(clpCode(0).Text) = 0 Then
        MsgBox "Document Type is required."
        clpCode(0).SetFocus
        Exit Function
    End If
    
    
    'Release 8.1
    'Send LostFocus on Document Type code so it is validated as per the Document Type Codes security
    flgWrongDocTypeCode = False
    Call clpCode_LostFocus(0)
    If flgWrongDocTypeCode = True Then
        clpCode(0).SetFocus
        Exit Function
    End If
    
        
'    'Release 8.1
'    Dim xTemplate As String
'    Dim SQLQ As String
'    Dim rs As New ADODB.Recordset
'
'    '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
'    xTemplate = ""
'    xTemplate = Get_Template(glbUserID)
'
'    If xTemplate = "" Or xTemplate = "TEMPLATE" Then
'        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
'    Else
'        '????Ticket #24808 -  Retrieve template's security profile
'        SQLQ = "SELECT MAINTAINABLE from HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
'    End If
'    'SQLQ = "SELECT ACCESSABLE from HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & glbUserID & "'"
'    SQLQ = SQLQ & " AND CODENAME='" & clpCode(0).Text & "'"
'    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
'    If rs.EOF = False And rs.BOF = False Then
'        If rs("MAINTAINABLE") = 0 Then
'        'If rs("ACCESSABLE") = 0 Then
'            MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(0).Text & "' Document Type Code.", vbOKOnly + vbInformation, "Authorization failed"
'            rs.Close
'            Set rs = Nothing
'            clpCode(0).SetFocus
'            Exit Function
'        End If
'    Else
'        MsgBox "You do not have Authority to 'Maintain' on '" & clpCode(0).Text & "' Document Type Code.", vbOKOnly + vbInformation, "Authorization failed"
'        rs.Close
'        Set rs = Nothing
'        clpCode(0).SetFocus
'        Exit Function
'    End If
'    rs.Close
'    Set rs = Nothing
    

chkDoc = True

Exit Function

chkDoc_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkDoc", "Document Attachment", "edit/Add")
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

Public Function IfExist()
    Dim SQLQ
    Dim rsTemp As New ADODB.Recordset
    Dim fldPrx As String
    
    IfExist = False
    fldPrx = ""
    
    If (Len(glbDocKey) = 0 And Len(glbJob) = 0) And (Not glbDocName = "Resume") And (Not glbDocName = "Termination") And (Not glbDocName = "OtherInfo") Then    'Release 8.1
        Exit Function
    ElseIf (Len(glbDocKey) = 0 Or glbDocKey = 0) And (glbDocName <> "Resume") And (glbDocName <> "Termination") And (glbDocName <> "Offer") And (glbDocName <> "Jobdescription") And (glbDocName <> "INCIDENT") And (glbDocName <> "EmployeeFlag") And (glbDocName <> "PositionSkill") And (glbDocName <> "OtherInfo") Then
        Exit Function
    End If
    
    Select Case glbDocName
        Case "Resume"
            SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(glbDocName) & "' AND RE_EMPNBR=" & glbLEE_ID
            'RsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            fldPrx = "RE"
        Case "Offer"
            SQLQ = "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND DJ_JOB= '" & glbJob & "'"
            SQLQ = SQLQ & " AND DJ_SDATE= " & Date_SQL(glbSDate)
            fldPrx = "DJ"
        Case "Jobdescription"
            SQLQ = "SELECT * FROM HRDOC_JOB WHERE DB_TYPE='" & UCase(glbDocName) & "'"
            SQLQ = SQLQ & " AND DB_JOB= '" & glbPos & "'"
            fldPrx = "DB"
        Case "Counsel"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR=" & glbLEE_ID
                'SQLQ = SQLQ & " AND DC_CLTYPE= '" & glbCounselType & "'"
                'SQLQ = SQLQ & " AND DC_COUDATE= " & Date_SQL(glbCounselDate)
            End If
            SQLQ = SQLQ & " AND DC_DOCKEY= " & glbDocKey & ""
            fldPrx = "DC"
        Case "Comments"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbLEE_ID
                'SQLQ = SQLQ & " AND DO_COTYPE= '" & glbCommentType & "'"
                'SQLQ = SQLQ & " AND DO_EDATE= " & Date_SQL(glbCommentDate)
            End If
            SQLQ = SQLQ & " AND DO_DOCKEY= " & glbDocKey & ""
            fldPrx = "DO"
        Case "INCIDENT"
            SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND DE_CASE= '" & glbJob & "'"
            SQLQ = SQLQ & " AND DE_DOCNO= '" & glbDocTmp & "'"
            fldPrx = "DE"
        Case "INJURYWF7"
            SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='" & UCase(glbDocName) & "' AND W7_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND W7_CASE= '" & glbJob & "'"
            SQLQ = SQLQ & " AND W7_DOCKEY= '" & glbDocKey & "'"
            fldPrx = "W7"
        Case "INJURYWF7_WRITTENOFR"
            SQLQ = "SELECT * FROM HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='" & UCase(glbDocName) & "' AND F7_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND F7_CASE= '" & glbJob & "'"
            SQLQ = SQLQ & " AND F7_DOCKEY= '" & glbDocKey & "'"
            fldPrx = "F7"
        Case "Performance"
            SQLQ = "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(glbDocName) & "' AND DH_EMPNBR=" & glbLEE_ID
            'SQLQ = SQLQ & " AND DH_JOB= '" & glbJob & "'"
            'SQLQ = SQLQ & " AND DH_PREVDATE= " & Date_SQL(glbSDate)
            SQLQ = SQLQ & " AND DH_DOCKEY= " & glbDocKey & ""
            fldPrx = "DH"
        Case "EdSem"
            SQLQ = "SELECT * FROM HRDOC_EDSEM WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND ES_DOCKEY= " & glbDocKey & ""
            fldPrx = "ES"
        Case "EdSem_Retest"
            SQLQ = "SELECT * FROM HRDOC_EDSEM_RETEST WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND ES_DOCKEY= " & glbDocKey & ""
            fldPrx = "ES"
        Case "FormalEdu"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR=" & glbLEE_ID
            End If
            SQLQ = SQLQ & " AND EU_DOCKEY= " & glbDocKey & ""
            fldPrx = "EU"
        Case "DollarEnt"
            SQLQ = "SELECT * FROM HRDOC_HRDOLENT WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID
            SQLQ = SQLQ & " AND DE_DOCKEY= " & glbDocKey & ""
            fldPrx = "DE"
        Case "Associations"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR=" & glbLEE_ID
            End If
            SQLQ = SQLQ & " AND TD_DOCKEY= " & glbDocKey & ""
            fldPrx = "TD"
        Case "Attendance"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR=" & glbLEE_ID
            End If
            SQLQ = SQLQ & " AND AD_DOCKEY= " & glbDocKey & ""
            fldPrx = "AD"
        Case "Termination"
            SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='" & UCase(glbDocName) & "' AND RE_EMPNBR=" & glbLEE_ID
            fldPrx = "RE"
        Case "EmployeeFlag"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(glbDocName) & "' AND EF_EMPNBR=" & glbLEE_ID
            End If
            fldPrx = "EF"
        Case "PositionSkill"
            SQLQ = "SELECT * FROM HRDOC_JOBSKL WHERE DS_TYPE='" & UCase(glbDocName) & "'"
            SQLQ = SQLQ & " AND DS_JOB= '" & glbPos & "'"
            SQLQ = SQLQ & " AND DS_SKILL= '" & glbPosSkill & "'"
            fldPrx = "DS"
        Case "OtherInfo"
            If glbtermopen Then
                SQLQ = "SELECT * FROM Term_HRDOC_HREMP_OTHER  WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            Else
                SQLQ = "SELECT * FROM HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbLEE_ID
            End If
            fldPrx = "ER"
        Case "LOA"
            'If glbtermopen Then
            '    SQLQ = "SELECT * FROM Term_HRDOC_HREMP_OTHER  WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbTERM_ID & " AND TERM_SEQ = " & glbTERM_Seq
            'Else
                SQLQ = "SELECT * FROM HRDOC_HRSTATUS WHERE SC_TYPE='" & UCase(glbDocName) & "' AND SC_EMPNBR=" & glbLEE_ID
            'End If
            SQLQ = SQLQ & " AND SC_DOCKEY= " & glbDocKey & ""
            fldPrx = "SC"
    End Select
    
    rsTemp.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    If Not rsTemp.EOF Then
        If glbDocName = "INCIDENT" Then
            If Not IsNull(rsTemp("DE_FILEEXT")) Then IfExist = True
        Else
            IfExist = True
        End If
        If Not IsNull(rsTemp(fldPrx & "_DOCTYPE")) Then
            clpCode(0).Text = rsTemp(fldPrx & "_DOCTYPE")
            txtDocDesc.Text = IIf(IsNull(rsTemp(fldPrx & "_USRDESC")), "", rsTemp(fldPrx & "_USRDESC"))
        End If
    End If
    rsTemp.Close

End Function
