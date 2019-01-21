VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmUEmailLoad 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Email Load"
   ClientHeight    =   7290
   ClientLeft      =   2565
   ClientTop       =   -1140
   ClientWidth     =   11400
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
   ScaleHeight     =   7290
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optEmail 
      Caption         =   "Load Email Setup"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2460
      Locked          =   -1  'True
      MaxLength       =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "00-File Name to import"
      Top             =   960
      Width           =   6450
   End
   Begin VB.CommandButton cmdEmailImp 
      Appearance      =   0  'Flat
      Caption         =   "Import"
      Height          =   280
      Left            =   9360
      TabIndex        =   4
      Tag             =   "Import the File"
      Top             =   960
      Width           =   855
   End
   Begin VB.CommandButton cmdEmailImpFile 
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
      Left            =   8955
      TabIndex        =   3
      Tag             =   "Select File to Import"
      Top             =   960
      Width           =   375
   End
   Begin MSComDlg.CommonDialog AttachmentDialog 
      Left            =   9720
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.OptionButton optEmail 
      Caption         =   "Update on Status/Dates screen"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   3070
   End
   Begin VB.Image imgHelp 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   3360
      Picture         =   "fuEmailLoad.frx":0000
      Stretch         =   -1  'True
      Top             =   480
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address Import File"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1005
      Width           =   2145
   End
   Begin VB.Image imgHelp1 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   6720
      Picture         =   "fuEmailLoad.frx":0442
      Stretch         =   -1  'True
      Top             =   480
      Width           =   270
   End
End
Attribute VB_Name = "frmUEmailLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Title$, msg$, DgDef As Variant, Response%

On Error GoTo Mod_Err
'If Not gSec_Upd_Email_Load Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If

Call cmdEmailImp_Click

Exit Sub

Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
     RollBack
    Resume Next
Else
    Unload Me
End If
End Sub

'Private Function CR_SnapEntitle()
'Dim SQLQ As String, SQLQ1 As String
'Dim snapMultiEmp As New ADODB.Recordset
'
'CR_SnapEntitle = False
'On Error GoTo CR_SnapEntitle_Err
'
'Screen.MousePointer = HOURGLASS
'
'Call getWSQLQ
'
''Jaddy Changed. Cound not join the position table it will cost so many problems.
'
'SQLQ = "SELECT ED_EMPNBR,ED_VACPC,ED_PVAC,ED_VAC,ED_PSICK,ED_SICK,ED_SICKT,ED_VACT, ED_ANNSICK, ED_ANNVAC, "
'SQLQ = SQLQ & " ED_DIV,ED_PT, "
'SQLQ = SQLQ & " ED_DEPTNO,ED_ORG, " 'NEW BY Frank for County of Elgin ticket #4653
'SQLQ = SQLQ & " ED_LOC,ED_SECTION,ED_SALDIST, "  'NEW
'SQLQ = SQLQ & " ED_EFDATE,ED_ETDATE,ED_EFDATES,ED_ETDATES, "  'NEW
'SQLQ = SQLQ & " ED_DOH, ED_SENDTE,ED_UNION,ED_LTHIRE,ED_USRDAT1,ED_LUSER,ED_LDATE,ED_LTIME"
'SQLQ = SQLQ & " FROM HREMP "
'SQLQ = SQLQ & " WHERE " & fglbESQLQ
'If snapEntitle.State <> 0 Then snapEntitle.Close
'If glbOracle Then
'    snapEntitle.CursorLocation = adUseServer
'End If
'snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic  'adLockPessimistic ''
'
'CR_SnapEntitle = True
'
'Exit Function
'
'CR_SnapEntitle_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CR_SnapEntitle", "Email Load", "Select")
'
'If gintRollBack% = False Then
'    Resume Next
'Else
'    Unload Me
'End If
'
'End Function

Private Sub cmdEmailImp_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE

glbOnTop = "FRMUEMAILLOAD"

End Sub

Private Sub Form_Load()

glbOnTop = "FRMUEMAILLOAD"

Call INI_Controls(Me)

Screen.MousePointer = DEFAULT

End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select FROM the menu the appropriate function."
glbUEnt = 0
Set frmUEmailLoad = Nothing  'carmen apr 2000
End Sub

'Private Sub getWSQLQ()
'
'fglbESQLQ = glbSeleDeptUn
'
'End Sub

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
    TF = True
    UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
End Sub

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = GetMassUpdateSecurities("EmailLoad_MassUpdate", glbUserID)
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

Private Sub cmdEmailImpFile_Click()
    glbDocName = "EmailSetup"
    
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

Private Sub cmdEmailImp_Click()
    Dim DgDef, Title$, msg$, Response%
    
    If Trim(txtFileName.Text) = "" Then
        MsgBox "File to import not selected. Please select the file to import.", vbExclamation
        cmdEmailImp.SetFocus
        Exit Sub
    ElseIf Dir(txtFileName.Text) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & txtFileName.Text & "]", vbExclamation
        cmdEmailImp.SetFocus
        Exit Sub
    Else
        Title$ = "Email Import"
        DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
        If optEmail(0) Then
            msg$ = "Are you sure you want to import this Email file?"
        Else
            msg$ = "Are you sure you want to import this Email Setup file?"
        End If
        Response% = MsgBox(msg$, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then    ' Evaluate response
            Exit Sub
        End If
        
        'Load Email Addresses
        If optEmail(0) Then
            Call Load_Email
        ElseIf optEmail(1) Then
            Call Load_EmailSetup
        End If
    End If

End Sub

Private Sub imgHelp_Click()
Dim MsgStr As String
    MsgStr = "Import File must be an Excel Spreadsheet with the following format: "
    MsgStr = MsgStr & Chr(10) & "        1. First row is a Header row."
    MsgStr = MsgStr & Chr(10) & "        2. Data to import must start from 2nd row."
    MsgStr = MsgStr & Chr(10) & "        3. Column order to Import:"
    MsgStr = MsgStr & Chr(10) & vbTab & "a. Column 1: Employee #"
    MsgStr = MsgStr & Chr(10) & vbTab & "b. Column 2: Email Address"
    MsgBox MsgStr, vbInformation, "info:HR - Import File Format"
End Sub

Private Sub Load_Email()
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim rsEMP As New adodb.Recordset
    Dim xSkipped As String
    Dim SQLQ As String
    Dim xEmail As String
    Dim xNum As Integer
    Dim xRows As Long
    Dim xRow As Long
    Dim xEmpnbr
    Dim xloaded As Long
    
    
    On Error GoTo Email_Err

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

    xloaded = 0
    
    For xRow = 2 To xRows
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
     
        xEmpnbr = exSheet.Cells(xRow, 1)
        xEmail = exSheet.Cells(xRow, 2)
        
        If Not IsNumeric(xEmpnbr) Or xEmpnbr = 0 Or Trim(xEmail) = "" Then
            xSkipped = xSkipped & xEmpnbr & "; "
            xNum = xNum + 1
            If xNum = 10 Then
                xSkipped = xSkipped & vbCrLf
                xNum = 0
            End If
        Else
            Set rsEMP = Nothing
            rsEMP.Open "SELECT ED_EMPNBR, ED_EMAIL FROM HREMP WHERE ED_EMPNBR =" & xEmpnbr, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsEMP.EOF Then
                rsEMP("ED_EMAIL") = Left(exSheet.Cells(xRow, 2), 60)
                rsEMP.Update
                
                xloaded = xloaded + 1
            Else
                xSkipped = xSkipped & xEmpnbr & "; "
                xNum = xNum + 1
                If xNum = 10 Then
                    xSkipped = xSkipped & vbCrLf
                    xNum = 0
                End If
            End If
            rsEMP.Close
            Set rsEMP = Nothing
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
        If xloaded > 0 Then
            MsgBox xloaded & " Employee's Email Addresses have been loaded successfully on Status/Dates screen.", vbOKOnly + vbInformation, "Import Email Addresses"
        Else
            MsgBox xloaded & " Employee's Email Addresses have been loaded on Status/Dates screen.", vbOKOnly + vbInformation, "Import Email Addresses"
        End If
    End If

Exit Sub

Email_Err:
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


Private Sub Load_EmailSetup()
    Dim exApp As Object, exBook As Object, exSheet As Object
    Dim rsEMP As New adodb.Recordset
    Dim rsEmail As New adodb.Recordset
    Dim rsSecure As New adodb.Recordset
    Dim xSkipped As String
    Dim SQLQ As String
    Dim xEmail, xServer, xUserName, xPassword, xSup As String
    Dim xSuper As Integer
    Dim xNum As Integer
    Dim xRows As Long
    Dim xRow As Long
    Dim xEmpnbr
    Dim xUserID As String
    
    
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
        
        'Jerry said to use User ID
        'xEmpnbr = exSheet.Cells(xRow, 1)
        xUserID = exSheet.Cells(xRow, 1) 'User ID
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

        If Trim(xUserID) = "" Or Trim(xEmail) = "" Or Trim(xServer) = "" Then
        'If Not IsNumeric(xEmpnbr) Or xEmpnbr = 0 Or Trim(xEmail) = "" Or Trim(xServer) = "" Then
            'xSkipped = xSkipped & xEmpnbr & "; "
            xSkipped = xSkipped & xUserID & "; "
            xNum = xNum + 1
            If xNum = 10 Then
                xSkipped = xSkipped & vbCrLf
                xNum = 0
            End If
        Else
            'Check if Security Profile exists for this User ID
            'Get the User ID
            'SQLQ = "SELECT USERID, EMPNBR FROM HR_SECURE_BASIC WHERE USERID = " & xEmpnbr
            SQLQ = "SELECT USERID, EMPNBR FROM HR_SECURE_BASIC WHERE USERID = '" & Replace(xUserID, "'", "''") & "'"
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
                'xSkipped = xSkipped & xEmpnbr & "; "
                xSkipped = xSkipped & xUserID & "; "
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
        'MsgBox "The Email address for the following Employee(s) have been skipped:" & vbCrLf & xSkipped, vbOKOnly + vbInformation, "Import Email Addresses"
        MsgBox "The Email address for the following User ID(s) have been skipped:" & vbCrLf & xSkipped, vbOKOnly + vbInformation, "Import Email Addresses"
    Else
        'MsgBox "Employee's Email Addresses have been loaded successfully on Email Setup screen.", vbOKOnly + vbInformation, "Import Email Addresses"
        MsgBox "User's Email Addresses have been loaded successfully on Email Setup screen.", vbOKOnly + vbInformation, "Import Email Addresses"
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

Private Sub imgHelp1_Click()
Dim MsgStr As String
    MsgStr = "Import File must be an Excel Spreadsheet with the following format: "
    MsgStr = MsgStr & Chr(10) & "        1. First row is a Header row."
    MsgStr = MsgStr & Chr(10) & "        2. Data to import must start from 2nd row."
    MsgStr = MsgStr & Chr(10) & "        3. Column order to Import:"
    MsgStr = MsgStr & Chr(10) & vbTab & "a. Column 1: User ID"
    MsgStr = MsgStr & Chr(10) & vbTab & "b. Column 2: Email Address"
    MsgStr = MsgStr & Chr(10) & vbTab & "c. Column 3: SMTP Server"
    MsgStr = MsgStr & Chr(10) & vbTab & "d. Column 4: SMTP Username"
    MsgStr = MsgStr & Chr(10) & vbTab & "e. Column 5: SMTP Password"
    MsgStr = MsgStr & Chr(10) & vbTab & "f. Column 6: Supervisor ('Y' for Yes or 'N' for No)"
    MsgBox MsgStr, vbInformation, "info:HR - Import File Format"
End Sub
