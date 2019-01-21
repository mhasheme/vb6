VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmSecuCopy 
   Caption         =   "Copy Security Settings"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtVerifyPwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   6
      Tag             =   "01- You must enter password"
      Top             =   1920
      Width           =   1425
   End
   Begin VB.TextBox txtFromUserID 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Tag             =   "01-You must Enter User ID"
      Top             =   120
      Width           =   1305
   End
   Begin VB.TextBox txtSecPwd 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4200
      PasswordChar    =   "*"
      TabIndex        =   14
      Tag             =   "01- You must enter password"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   5
      Tag             =   "01- You must enter password"
      Top             =   1560
      Width           =   1425
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   50
      TabIndex        =   4
      Tag             =   "01- You must enter User Name"
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtCopyUserID 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      MaxLength       =   25
      TabIndex        =   2
      Tag             =   "01-You must Enter User ID"
      Top             =   480
      Width           =   1305
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   2550
      Width           =   5760
      _Version        =   65536
      _ExtentX        =   10160
      _ExtentY        =   979
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
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
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
         Left            =   1496
         TabIndex        =   9
         Tag             =   "Save changes made"
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         Left            =   3169
         TabIndex        =   8
         Top             =   30
         Width           =   1095
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8490
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   3
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin INFOHR_Controls.EmployeeLookup elpEmpLookup 
      Height          =   285
      Left            =   1725
      TabIndex        =   3
      Tag             =   "11-Employee Number"
      Top             =   840
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   503
      RefreshDescriptionWhen=   2
   End
   Begin INFOHR_Controls.CodeLookup clpSalDist 
      Height          =   285
      Left            =   1725
      TabIndex        =   7
      Top             =   2280
      Visible         =   0   'False
      Width           =   4395
      _ExtentX        =   7752
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   6
      LookupType      =   8
   End
   Begin VB.Label lblSalDist 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Distribution"
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   2325
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Verify Password"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   16
      Top             =   1965
      Width           =   1125
   End
   Begin VB.Label lblCopyUesrID 
      AutoSize        =   -1  'True
      Caption         =   "From User ID"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   165
      Width           =   930
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1605
      Width           =   825
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1245
      Width           =   495
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Employee#"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   885
      Width           =   795
   End
   Begin VB.Label lblCopyUesrID 
      AutoSize        =   -1  'True
      Caption         =   "To User ID"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   525
      Width           =   780
   End
End
Attribute VB_Name = "frmSecuCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xIsCopyByPayID As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim X, sUSERID, sCompPass, ICompPass
Dim sql As String
Dim OPwd As String
Dim rsBASIC As New ADODB.Recordset
Dim rsBASICTemp As New ADODB.Recordset
Dim flgExpDate As Boolean
    
If xIsCopyByPayID Then 'Family Day Ticket #24729 01/20/2014
    Call UptCopyByPayID
    Exit Sub
End If

If Not CriCheck() Then
    Exit Sub
End If

On Error GoTo Err_ExpiryDate

    ''*** CHECK TO SEE IF EMPLOYEE EXIST IN HR_SECURE_BASIC
    sql = "SELECT * FROM HR_SECURE_BASIC WHERE USERID= '" & Replace(txtCopyUserID, "'", "''") & "'"
    rsBASIC.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsBASIC.EOF Then
        MsgBox "This User ID " & txtCopyUserID & " exists already!"
        rsBASIC.Close
        Exit Sub
    End If
    
    'Ticket #16459 WorkSafe NB (whscc) Frank 03/20/2009
    'txtSecPwd.Text = EncryptPassword(txtPassword.Text)
    If gsMultiLang = "YES" Then 'whscc
        txtSecPwd.Text = EncryptPasswordMultiLang(txtPassword.Text)
    Else
        txtSecPwd.Text = EncryptPassword(txtPassword.Text)
    End If
    
 '*** Get info for selected Employee from Grid in Basic Table to copy the optional values
              
    sql = "SELECT * FROM HR_SECURE_BASIC WHERE USERID= '" & Replace(txtFromUserID, "'", "''") & "'"
    rsBASICTemp.Open sql, gdbAdoIhr001, adOpenStatic
    If rsBASICTemp.EOF Then
        Exit Sub
    End If
    rsBASIC.AddNew
    If IsNumeric(elpEmpLookup.Text) Then rsBASIC("EMPNBR") = elpEmpLookup.Text
    rsBASIC("COMPNO") = "001"
    rsBASIC("USERID") = txtCopyUserID.Text
    'Ticket #29350
    'rsBASIC("USERNAME") = Replace(txtUserName.Text, "'", "'+chr(39)+'")
    rsBASIC("USERNAME") = txtUserName.Text
    rsBASIC("Password") = txtSecPwd.Text
    rsBASIC("Empnbr_Based") = rsBASICTemp("Empnbr_Based")
    rsBASIC("COUNTRY") = rsBASICTemp("COUNTRY")
    rsBASIC("LDATE") = Date
    rsBASIC("LTIME") = Time$
    rsBASIC("LUSER") = glbUserID
    
    'Ticket #20585 - if there is a template then update with Template Name
    If Not IsNull(rsBASICTemp("SECURE_TEMPLATE")) Then
        If rsBASICTemp("SECURE_TEMPLATE") = "TEMPLATE" Then
            rsBASIC("SECURE_TEMPLATE") = txtFromUserID.Text 'Template Name
        Else
            rsBASIC("SECURE_TEMPLATE") = rsBASICTemp("SECURE_TEMPLATE") 'From User's Template to New User
        End If
    End If
    
    'Ticket #22077 - Copy over the Expiry Days and compute Expiration Date.
    rsBASIC("PS_EXPIR_DAYS") = IIf(IsNull(rsBASICTemp("PS_EXPIR_DAYS")), 0, rsBASICTemp("PS_EXPIR_DAYS"))
    If IsNumeric(rsBASICTemp("PS_EXPIR_DAYS")) Then
        If rsBASICTemp("PS_EXPIR_DAYS") <> 0 Then
            'Ticket #22893 - If the date is not valid then update with Upper limit
            'rsBASIC("PS_EXPIR_DATE") = DateAdd("d", rsBASICTemp("PS_EXPIR_DAYS"), Date)
            flgExpDate = True
            rsBASIC("PS_EXPIR_DATE") = IIf(IsDate(DateAdd("d", rsBASICTemp("PS_EXPIR_DAYS"), Date)), DateAdd("d", rsBASICTemp("PS_EXPIR_DAYS"), Date), CVDate(Format("12/31/9999", "mm/dd/yyyy")))
            flgExpDate = False
        End If
    End If
    
    'Ticket #22293 - if there is a Timesheet Template attached then copy over
    If Not IsNull(rsBASICTemp("TS_TPID")) Then
        rsBASIC("TS_TPID") = rsBASICTemp("TS_TPID")
    End If
    
    rsBASIC.Update
       
    Call CopySecurity
                    
    Unload Me
Exit Sub
Err_ExpiryDate:
    If Err.Number = 13 And flgExpDate Then
        rsBASIC("PS_EXPIR_DATE") = CVDate(Format("12/31/9999", "mm/dd/yyyy"))
        Resume Next
    End If
End Sub

Private Sub CopySecurity()
Dim X, sUSERID
Dim sql As String
Dim rsACCESS As New ADODB.Recordset
Dim rsINSERT As New ADODB.Recordset
Dim rsCopySecurity As New ADODB.Recordset
Dim xTemplate As String

    '????Ticket #24808 - If copied from User is Templated based then the New Copied User's Profile will only be updated
    'up to Employee # Based Security, Password Expiry and Department Security

    '????Ticket #24808 -  Get From User's Template if there is one to retrieve template's Department security profile
    xTemplate = ""
    xTemplate = Get_Template(txtFromUserID)
    
    '????Ticket #24808 - Only copy if normal User
    If xTemplate = "" Then
        sql = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
        rsCopySecurity.Open sql, gdbAdoIhr001, adOpenForwardOnly
        
        If rsCopySecurity.EOF Then
            MsgBox "This User has no security record to copy", vbOKOnly, "Error finding Employee"
            Exit Sub
        Else
            sql = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
            rsACCESS.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        End If
        
        MDIMain.panHelp(0).Caption = "Please wait while system copies security..."
        If Not rsACCESS.EOF Then
            ''*** if employee exist in security access then delete them
            Screen.MousePointer = HOURGLASS
            Do While Not rsACCESS.EOF
                rsACCESS.Delete
                rsACCESS.MoveNext
            Loop
        End If

    '    ''*** if employee does not exist in access security then copy security starts
        Do While Not rsCopySecurity.EOF
             rsACCESS.AddNew
             rsACCESS("USERID") = txtCopyUserID.Text
             rsACCESS("FUNCTION") = rsCopySecurity("FUNCTION")
             rsACCESS("ACCESSABLE") = rsCopySecurity("ACCESSABLE")
             rsACCESS("Maintainable") = rsCopySecurity("Maintainable")
             rsACCESS("CODENAME") = rsCopySecurity("CODENAME")
             rsACCESS("LDATE") = Date
             rsACCESS("LTIME") = Time$
             rsACCESS("LUSER") = glbUserID
    
             rsACCESS.Update
             rsCopySecurity.MoveNext
        Loop
        rsACCESS.Close
        rsCopySecurity.Close
    End If
    
    
    'Copying Department Security
    Dim rsFrmSecDept As New ADODB.Recordset
    Dim rsToSecDept As New ADODB.Recordset
    
    sql = "SELECT * FROM HRPASDEP WHERE PD_USERID='" & Replace(txtFromUserID, "'", "''") & "'"
    rsFrmSecDept.Open sql, gdbAdoIhr001, adOpenForwardOnly
    
    sql = "SELECT * FROM HRPASDEP WHERE PD_USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
    rsToSecDept.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    'If the record is already existing then delete it.
    Do While Not rsToSecDept.EOF
        rsToSecDept.Delete
        rsToSecDept.MoveNext
    Loop

    'Add new Dept Security
    Do While Not rsFrmSecDept.EOF
        rsToSecDept.AddNew
        rsToSecDept("PD_COMPNO") = "001"
        rsToSecDept("PD_USERID") = txtCopyUserID.Text
        rsToSecDept("PD_DEPT") = rsFrmSecDept("PD_DEPT")
        rsToSecDept("PD_ORG") = rsFrmSecDept("PD_ORG")
        rsToSecDept("PD_DIV") = rsFrmSecDept("PD_DIV")
        rsToSecDept("PD_SECTION") = rsFrmSecDept("PD_SECTION")
        rsToSecDept("PD_ADMINBY") = rsFrmSecDept("PD_ADMINBY")
        'Ticket #22682 - Release 8.0
        rsToSecDept("PD_LOC") = rsFrmSecDept("PD_LOC")
        rsToSecDept("PD_REGION") = rsFrmSecDept("PD_REGION")
        
        rsToSecDept("PD_INCLEMPNBR") = rsFrmSecDept("PD_INCLEMPNBR")
        rsToSecDept("PD_EXCLEMPNBR") = rsFrmSecDept("PD_EXCLEMPNBR")
        rsToSecDept.Update
        
        rsFrmSecDept.MoveNext
    Loop
    rsToSecDept.Close
    rsFrmSecDept.Close

    
    'Ticket #30508 - Applicant Tracking Enhancement
    'Copying Requisition Security
    Dim rsFrmSecRequi As New ADODB.Recordset
    Dim rsToSecRequi As New ADODB.Recordset
    
    sql = "SELECT * FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
    rsFrmSecRequi.Open sql, gdbAdoIhr001, adOpenForwardOnly
    
    sql = "SELECT * FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
    rsToSecRequi.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    'If the record is already existing then delete it.
    Do While Not rsToSecRequi.EOF
        rsToSecRequi.Delete
        rsToSecRequi.MoveNext
    Loop

    'Add new Requisition Security
    Do While Not rsFrmSecRequi.EOF
        rsToSecRequi.AddNew
        rsToSecRequi("COMPNO") = "001"
        rsToSecRequi("USERID") = txtCopyUserID.Text
        rsToSecRequi("RS_POSTYPE") = rsFrmSecRequi("RS_POSTYPE")
        rsToSecRequi("RS_ORG") = rsFrmSecRequi("RS_ORG")
        rsToSecRequi("RS_GRPCD") = rsFrmSecRequi("RS_GRPCD")
        rsToSecRequi("RS_STATUS") = rsFrmSecRequi("RS_STATUS")
        
        rsToSecRequi("RS_INCLJOB") = rsFrmSecRequi("RS_INCLJOB")
        rsToSecRequi("RS_EXCLJOB") = rsFrmSecRequi("RS_EXCLJOB")
        
        rsToSecRequi("LDATE") = Date
        rsToSecRequi("LTIME") = Time$
        rsToSecRequi("LUSER") = glbUserID
        rsToSecRequi.Update
        
        rsFrmSecRequi.MoveNext
    Loop
    rsToSecRequi.Close
    rsFrmSecRequi.Close

    
    '????Ticket #24808 - Only copy if normal User
    If xTemplate = "" Then
        'Copying Comments Security
        Dim rsFrmSecComments As New ADODB.Recordset
        Dim rsToSecComments As New ADODB.Recordset
        
        sql = "SELECT * FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
        rsFrmSecComments.Open sql, gdbAdoIhr001, adOpenForwardOnly
        
        sql = "SELECT * FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
        rsToSecComments.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        'If the record is already existing then delete it.
        Do While Not rsToSecComments.EOF
            rsToSecComments.Delete
            rsToSecComments.MoveNext
        Loop
        
        Do While Not rsFrmSecComments.EOF
            rsToSecComments.AddNew
            rsToSecComments("COMPNO") = "001"
            rsToSecComments("USERID") = txtCopyUserID.Text
            rsToSecComments("ACCESSABLE") = rsFrmSecComments("ACCESSABLE")
            rsToSecComments("MAINTAINABLE") = rsFrmSecComments("MAINTAINABLE")
            rsToSecComments("CODENAME") = rsFrmSecComments("CODENAME")
            rsToSecComments("DESCRIPTION") = rsFrmSecComments("DESCRIPTION")
            rsToSecComments("LDATE") = Date
            rsToSecComments("LTIME") = Time$
            rsToSecComments("LUSER") = glbUserID
            rsToSecComments.Update
            
            rsFrmSecComments.MoveNext
        Loop
        rsToSecComments.Close
        rsFrmSecComments.Close
        
        
        'Copying Custom Features Security
        Dim rsFrmSecCustmRpt As New ADODB.Recordset
        Dim rsToSecCustmRpt As New ADODB.Recordset
        
        sql = "SELECT * FROM HR_SECRPT WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
        rsFrmSecCustmRpt.Open sql, gdbAdoIhr001, adOpenForwardOnly
        
        sql = "SELECT * FROM HR_SECRPT WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
        rsToSecCustmRpt.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        'If the record is already existing then delete it.
        Do While Not rsToSecCustmRpt.EOF
            rsToSecCustmRpt.Delete
            rsToSecCustmRpt.MoveNext
        Loop
        
        Do While Not rsFrmSecCustmRpt.EOF
            rsToSecCustmRpt.AddNew
            rsToSecCustmRpt("COMPNO") = "001"
            rsToSecCustmRpt("USERID") = txtCopyUserID.Text
            rsToSecCustmRpt("FUNCTION") = rsFrmSecCustmRpt("FUNCTION")
            rsToSecCustmRpt("ACCESSABLE") = rsFrmSecCustmRpt("ACCESSABLE")
            rsToSecCustmRpt("Maintainable") = rsFrmSecCustmRpt("Maintainable")
            rsToSecCustmRpt("CODENAME") = rsFrmSecCustmRpt("CODENAME")
            rsToSecCustmRpt("LDATE") = Date
            rsToSecCustmRpt("LTIME") = Time$
            rsToSecCustmRpt("LUSER") = glbUserID
            rsToSecCustmRpt.Update
            
            rsFrmSecCustmRpt.MoveNext
        Loop
        rsToSecCustmRpt.Close
        rsFrmSecCustmRpt.Close
        
        'Copying Linamar's Custom Features
        If glbLinamar Then
            Dim rsFrmSecCustmFeat As New ADODB.Recordset
            Dim rsToSecCustmFeat As New ADODB.Recordset
            
            sql = "SELECT * FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
            rsFrmSecCustmFeat.Open sql, gdbAdoIhr001, adOpenForwardOnly
            
            sql = "SELECT * FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
            rsToSecCustmFeat.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
            'If the record is already existing then delete it.
            Do While Not rsToSecCustmFeat.EOF
                rsToSecCustmFeat.Delete
                rsToSecCustmFeat.MoveNext
            Loop
            
            Do While Not rsFrmSecCustmFeat.EOF
                rsToSecCustmFeat.AddNew
                rsToSecCustmFeat("COMPNO") = "001"
                rsToSecCustmFeat("USERID") = txtCopyUserID.Text
                rsToSecCustmFeat("FUNCTION") = rsFrmSecCustmFeat("FUNCTION")
                rsToSecCustmFeat("ACCESSABLE") = rsFrmSecCustmFeat("ACCESSABLE")
                rsToSecCustmFeat("Maintainable") = rsFrmSecCustmFeat("Maintainable")
                rsToSecCustmFeat("CODENAME") = rsFrmSecCustmFeat("CODENAME")
                rsToSecCustmFeat("LDATE") = Date
                rsToSecCustmFeat("LTIME") = Time$
                rsToSecCustmFeat("LUSER") = glbUserID
                rsToSecCustmFeat.Update
                
                rsFrmSecCustmFeat.MoveNext
            Loop
            rsToSecCustmFeat.Close
            rsFrmSecCustmFeat.Close
        End If
        
        'Copying Follow Up Security
        Dim rsFrmSecFollowUp As New ADODB.Recordset
        Dim rsToSecFollowUp As New ADODB.Recordset
        
        sql = "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
        rsFrmSecFollowUp.Open sql, gdbAdoIhr001, adOpenForwardOnly
        
        sql = "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
        rsToSecFollowUp.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        'If the record is already existing then delete it.
        Do While Not rsToSecFollowUp.EOF
            rsToSecFollowUp.Delete
            rsToSecFollowUp.MoveNext
        Loop
        
        Do While Not rsFrmSecFollowUp.EOF
            rsToSecFollowUp.AddNew
            rsToSecFollowUp("COMPNO") = "001"
            rsToSecFollowUp("USERID") = txtCopyUserID.Text
            rsToSecFollowUp("ACCESSABLE") = rsFrmSecFollowUp("ACCESSABLE")
            rsToSecFollowUp("MAINTAINABLE") = rsFrmSecFollowUp("MAINTAINABLE")
            rsToSecFollowUp("CODENAME") = rsFrmSecFollowUp("CODENAME")
            rsToSecFollowUp("DESCRIPTION") = rsFrmSecFollowUp("DESCRIPTION")
            rsToSecFollowUp("LDATE") = Date
            rsToSecFollowUp("LTIME") = Time$
            rsToSecFollowUp("LUSER") = glbUserID
            rsToSecFollowUp.Update
            
            rsFrmSecFollowUp.MoveNext
        Loop
        rsToSecFollowUp.Close
        rsFrmSecFollowUp.Close
        
        'Copying Attendance Reason Code Security
        Dim rsFrmSecAttend As New ADODB.Recordset
        Dim rsToSecAttend As New ADODB.Recordset
        
        sql = "SELECT * FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
        rsFrmSecAttend.Open sql, gdbAdoIhr001, adOpenForwardOnly
        
        sql = "SELECT * FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
        rsToSecAttend.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        'If the record is already existing then delete it.
        Do While Not rsToSecAttend.EOF
            rsToSecAttend.Delete
            rsToSecAttend.MoveNext
        Loop
        
        Do While Not rsFrmSecAttend.EOF
            rsToSecAttend.AddNew
            rsToSecAttend("COMPNO") = "001"
            rsToSecAttend("USERID") = txtCopyUserID.Text
            rsToSecAttend("ACCESSABLE") = rsFrmSecAttend("ACCESSABLE")
            rsToSecAttend("MAINTAINABLE") = rsFrmSecAttend("MAINTAINABLE")
            rsToSecAttend("CODENAME") = rsFrmSecAttend("CODENAME")
            rsToSecAttend("DESCRIPTION") = rsFrmSecAttend("DESCRIPTION")
            rsToSecAttend("LDATE") = Date
            rsToSecAttend("LTIME") = Time$
            rsToSecAttend("LUSER") = glbUserID
            rsToSecAttend.Update
            
            rsFrmSecAttend.MoveNext
        Loop
        rsToSecAttend.Close
        rsFrmSecAttend.Close
    
        'Release 8.1
        'Copying Document Type Security
        Dim rsFrmSecDocType As New ADODB.Recordset
        Dim rsToSecDocType As New ADODB.Recordset
        
        sql = "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(txtFromUserID, "'", "''") & "'"
        rsFrmSecDocType.Open sql, gdbAdoIhr001, adOpenForwardOnly
        
        sql = "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(txtCopyUserID, "'", "''") & "'"
        rsToSecDocType.Open sql, gdbAdoIhr001, adOpenStatic, adLockPessimistic
        'If the record is already existing then delete it.
        Do While Not rsToSecDocType.EOF
            rsToSecDocType.Delete
            rsToSecDocType.MoveNext
        Loop
        
        Do While Not rsFrmSecDocType.EOF
            rsToSecDocType.AddNew
            rsToSecDocType("COMPNO") = "001"
            rsToSecDocType("USERID") = txtCopyUserID.Text
            rsToSecDocType("ACCESSABLE") = rsFrmSecDocType("ACCESSABLE")
            rsToSecDocType("MAINTAINABLE") = rsFrmSecDocType("MAINTAINABLE")
            rsToSecDocType("CODENAME") = rsFrmSecDocType("CODENAME")
            rsToSecDocType("DESCRIPTION") = rsFrmSecDocType("DESCRIPTION")
            rsToSecDocType("LDATE") = Date
            rsToSecDocType("LTIME") = Time$
            rsToSecDocType("LUSER") = glbUserID
            rsToSecDocType.Update
            
            rsFrmSecDocType.MoveNext
        Loop
        rsToSecDocType.Close
        rsFrmSecDocType.Close
    End If
    
    Screen.MousePointer = DEFAULT
    MDIMain.panHelp(0).Caption = "Security Copying Done"
    MsgBox "Security added for '" & txtCopyUserID & "' successfully.", vbInformation, "Security Added"

    Unload Me
End Sub

Private Function CriCheck()
Dim X%

CriCheck = False

'If Len(elpEmpLookup.Text) = 0 Then
'    MsgBox "User Name missed!", vbCritical, "Error Occurred while adding User Name"
'    elpEmpLookup.SetFocus
'    Exit Function
'End If

If Len(txtCopyUserID.Text) = 0 Then
    MsgBox "To UserID missed!", vbCritical, "Error Occurred while adding UserID"
    txtCopyUserID.SetFocus
    Exit Function
End If

If Len(txtUserName.Text) = 0 Then
    MsgBox "User Name missed!", vbCritical, "Error Occurred while adding User Name"
    txtUserName.SetFocus
    Exit Function
End If

'Ticket #24031 - Employee # required if Employee # Based Security
If frmSECURE.chkEESecurity.Value = True And Len(elpEmpLookup) = 0 Then
    MsgBox "'From User ID' has 'Employee Number Based Security' checked 'To User ID' must have 'Employee #'", vbOKOnly + vbExclamation, "info:HR Security"
    elpEmpLookup.SetFocus
    Exit Function
End If

If Len(txtPassword) < 1 Or Len(txtPassword) > 15 Then
    MsgBox "Invalid Password (must be between 1 and 15 characters)'"
    txtPassword.SetFocus
    Exit Function
End If
If Len(txtVerifyPwd) < 1 Or Len(txtVerifyPwd) > 15 Then
    MsgBox "Invalid Verify Password (must be between 1 and 15 characters)'"
    txtVerifyPwd.SetFocus
    Exit Function
End If
If Not (txtPassword = txtVerifyPwd) Then
    MsgBox "Password is not equal to Verify Password"
    txtVerifyPwd.SetFocus
    Exit Function
End If

CriCheck = True
End Function

Private Sub elpEmpLookup_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub elpEmpLookup_LostFocus()
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
If Len(elpEmpLookup) > 0 Then
    txtUserName = elpEmpLookup.Caption
    If Len(txtCopyUserID) = 0 Then
        txtCopyUserID = elpEmpLookup
    End If
End If

End Sub

Private Sub Form_Load()

If xIsCopyByPayID Then 'Family Day Ticket #24729 01/20/2014
    Me.Caption = "Copy to New Payroll ID"
    Call ScreensetForCopyPayID
End If

Call INI_Controls(Me)

'elpEmpLookup.SetFocus

End Sub

'Public Property let IsCopyByPayID() As Boolean
'
'End Property

Public Property Let IsCopyByPayID(vData As Boolean)
xIsCopyByPayID = vData
End Property

Private Sub ScreensetForCopyPayID() 'Family Day Ticket #24729 01/20/2014
Dim xOldCompCode As String
Dim xNewCompCode As String

    lblCopyUesrID(1).Caption = "From Payroll ID"
    lblCopyUesrID(0).Caption = "To Payroll ID"
    lblSalDist.Caption = lStr("Salary Distribution")
    lblSalDist.Top = lblTitle(3).Top
    clpSalDist.Top = elpEmpLookup.Top
    lblSalDist.Visible = True
    clpSalDist.Visible = True
    
    lblTitle(0).Visible = False
    lblTitle(1).Visible = False
    lblTitle(2).Visible = False
    lblTitle(3).Visible = False
    elpEmpLookup.Visible = False
    txtUserName.Visible = False
    txtPassword.Visible = False
    txtVerifyPwd.Visible = False
    
    xOldCompCode = GetEmpData(glbLEE_ID, "ED_SALDIST")
    If Len(xOldCompCode) > 0 Then
        If xOldCompCode = "1611" Then xNewCompCode = "0019"
        If xOldCompCode = "0019" Then xNewCompCode = "1611"
        clpSalDist.Text = xNewCompCode
    End If
End Sub

Private Sub UptCopyByPayID()
Dim SQLQ As String
Dim rsEmpp As New ADODB.Recordset
Dim rsPA As New ADODB.Recordset
Dim a As Integer, Msg As String
Dim xOldEmpNo As Long
Dim xNextEmpNo As Long
Dim xPayID As String
Dim xFList As String
Dim xOldCompCode As String
Dim xNewCompCode As String

    If Len(txtFromUserID.Text) = 0 Then
        MsgBox "From Payroll ID missed!", vbCritical, "Error Occurred"
        txtFromUserID.SetFocus
        Exit Sub
    End If
    If Len(txtCopyUserID.Text) = 0 Then
        MsgBox "To Payroll ID missed!", vbCritical, "Error Occurred"
        txtCopyUserID.SetFocus
        Exit Sub
    Else
        If Not Len(txtCopyUserID.Text) = 9 Then
            MsgBox "Payroll ID must be 9 digits", vbCritical, "Error Occurred"
            txtCopyUserID.SetFocus
            Exit Sub
        End If
    End If
    If Len(clpSalDist.Text) = 0 Then
        MsgBox lStr("Salary Distribution") & " missed!", vbCritical, "Error Occurred"
        clpSalDist.SetFocus
        Exit Sub
    Else
        If clpSalDist.Caption = "Unassigned" Then
            MsgBox lStr("Salary Distribution") & " must be valid"
            clpSalDist.SetFocus
            Exit Sub
        End If
    End If
    xNewCompCode = clpSalDist.Text
    
    xOldEmpNo = glbLEE_ID
    
    'check if Payroll ID exist
    xPayID = txtCopyUserID.Text
    SQLQ = "SELECT * FROM HREMP WHERE ED_PAYROLL_ID = '" & xPayID & "' "
    If rsEmpp.State <> 0 Then rsEmpp.Close
    rsEmpp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpp.EOF Then
        MsgBox "This Payroll ID '" & xPayID & "' exists already!"
        Exit Sub
    End If
    rsEmpp.Close
    
    xOldCompCode = ""
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xOldEmpNo & " "
    If rsEmpp.State <> 0 Then rsEmpp.Close
    rsEmpp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpp.EOF Then
        If Not IsNull(rsEmpp("ED_SALDIST")) Then
            xOldCompCode = rsEmpp("ED_SALDIST")
        End If
    End If
    rsEmpp.Close
    
    If xOldCompCode = xNewCompCode Then
        MsgBox "Cannot copy the employee in the same " & lStr("Salary Distribution") & " "
        Exit Sub
    End If
    
    Msg = "Are you sure you want to copy this Payroll ID? "
    a% = MsgBox(Msg, 36, "Confirm Copy")
    If a% <> 6 Then Exit Sub

    rsPA.Open "select PC_NEXT_AVAILABLE_NBR,PC_FEDTAX,PC_PROVTAX from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
    glbNextEmpl = rsPA("PC_NEXT_AVAILABLE_NBR")
    xNextEmpNo = CLng(glbNextEmpl)
    glbNextEmpl = glbNextEmpl + 1
    'If glbCompSerial = "S/N - 2241W" Then '  Granite Club
    '    Call Check_EMPLOYEE_Number(glbNextEmpl)
    'Else
        rsPA("PC_NEXT_AVAILABLE_NBR") = glbNextEmpl
        rsPA.Update
    'End If
    rsPA.Close
                
    'copy the data in hremp
    xFList = Get_Fields(gdbAdoIhr001, "HREMP", "ED_EMPNBR, ED_PAYROLL_ID,ED_SALDIST")
    SQLQ = "INSERT INTO HREMP (" & xFList & ",ED_EMPNBR,ED_PAYROLL_ID,ED_SALDIST) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ", "
    'SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & xNextEmpNo & " As ED_EMPNBR, "
    SQLQ = SQLQ & "'" & xPayID & "' " & " As ED_PAYROLL_ID, "
    SQLQ = SQLQ & "'" & xNewCompCode & "' " & " As ED_SALDIST "
    SQLQ = SQLQ & "FROM HREMP "
    SQLQ = SQLQ & "WHERE (HREMP.ED_EMPNBR = " & xOldEmpNo & " )"
    gdbAdoIhr001.Execute SQLQ

    'create audit for new employee
    Call NewPayIDHREMPAudit(xNextEmpNo)
    
    Call get_CopyToPayIDFormShort
    
    'call position and salary screen for new employee
    Call FamilyDayPositionSalary(xNextEmpNo, xPayID, xOldEmpNo)

    
    'MsgBox "   Finished!   "
    'Unload Me
    
End Sub


Private Function NewPayIDHREMPAudit(EEID&)
Dim SQLQ As String
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTC As New ADODB.Recordset
Dim xProvNbr, xADD, xPROV
Dim Langs 'George Apr 4,2006 #10574
'Dim TIHR_DB As Database

On Error GoTo NewPayIDHREMPAudit_ERR

NewPayIDHREMPAudit = False


rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

rsTC.Open "select * from HREMP where ED_EMPNBR = " & EEID&, gdbAdoIhr001, adOpenKeyset

If rsTC.EOF Then
    'MsgBox "SYSTEM ERROR - READING TERM_EMP"
    Exit Function
End If

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_COMPNO") = "001"
rsTA("AU_NEWEMP") = "Y"
rsTA("AU_TYPE") = "A"
rsTA("AU_DIV") = rsTC("ED_DIV")
rsTA("AU_DIVUPL") = rsTC("ED_DIV")
rsTA("AU_LOC") = rsTC("ED_LOC")
rsTA("AU_EMPNBR") = EEID 'rsTC("ED_EMPNBR") 'Ticket
rsTA("AU_TITLE") = rsTC("ED_TITLE")
rsTA("AU_SURNAME") = rsTC("ED_SURNAME")
rsTA("AU_FNAME") = rsTC("ED_FNAME")
rsTA("AU_ADDR1") = rsTC("ED_ADDR1")
rsTA("AU_ADDR2") = rsTC("ED_ADDR2")
rsTA("AU_CITY") = rsTC("ED_CITY")
rsTA("AU_PROV") = rsTC("ED_PROV")
rsTA("AU_PCODE") = rsTC("ED_PCODE")
rsTA("AU_PHONE") = rsTC("ED_PHONE")
rsTA("AU_PROVEMP") = rsTC("ED_PROVEMP")
rsTA("AU_PROVRES") = rsTC("ED_PROVEMP")
rsTA("AU_COUNTRY") = rsTC("ED_COUNTRY")
rsTA("AU_UIC") = rsTC("ED_UIC")
rsTA("AU_PENSION") = rsTC("ED_PENSION")
rsTA("AU_CPP") = rsTC("ED_CPP")
rsTA("AU_GROSSCD") = rsTC("ED_GROSSCD")
rsTA("AU_GARN") = rsTC("ED_GARN")
rsTA("AU_ELIGIBLE") = rsTC("ED_ELIGIBLE")
rsTA("AU_EARLYR") = rsTC("ED_EARLYR")
rsTA("AU_NORMALR") = rsTC("ED_NORMALR")
rsTA("AU_LATESTR") = rsTC("ED_LATESTR")
rsTA("AU_SIN") = rsTC("ED_SIN")
rsTA("AU_PT") = rsTC("ED_PT")
rsTA("AU_PTUPL") = rsTC("ED_PT")
rsTA("AU_EMPTYPE") = rsTC("ED_EMPTYPE")
rsTA("AU_SEX") = rsTC("ED_SEX")
If rsTC("ED_SMOKER") <> 0 Then
    rsTA("AU_SMOKER") = "Yes"
Else
    rsTA("AU_SMOKER") = "No"
End If
rsTA("AU_MSTAT") = rsTC("ED_MSTAT")
rsTA("AU_DEPTNO") = rsTC("ED_DEPTNO")
rsTA("AU_DOB") = rsTC("ED_DOB")
rsTA("AU_DOH") = rsTC("ED_DOH")
rsTA("AU_SENDTE") = rsTC("ED_SENDTE")
rsTA("AU_LTHIRE") = rsTC("ED_LTHIRE")
rsTA("AU_DEPOSIT") = rsTC("ED_DEPOSIT")
rsTA("AU_BRANCH") = rsTC("ED_BRANCH")
rsTA("AU_BANK") = rsTC("ED_BANK")
rsTA("AU_ACCOUNT") = rsTC("ED_ACCOUNT")
rsTA("AU_AMTDEPOSIT") = rsTC("ED_AMTDEPOSIT")
rsTA("AU_PCDEPOSIT") = rsTC("ED_PCDEPOSIT")
rsTA("AU_DEPOSIT2") = rsTC("ED_DEPOSIT2")
rsTA("AU_BRANCH2") = rsTC("ED_BRANCH2")
rsTA("AU_BANK2") = rsTC("ED_BANK2")
rsTA("AU_ACCOUNT2") = rsTC("ED_ACCOUNT2")
rsTA("AU_AMTDEPOSIT2") = rsTC("ED_AMTDEPOSIT2")
rsTA("AU_PCDEPOSIT2") = rsTC("ED_PCDEPOSIT2")
rsTA("AU_DEPOSIT3") = rsTC("ED_DEPOSIT3")
rsTA("AU_BRANCH3") = rsTC("ED_BRANCH3")
rsTA("AU_BANK3") = rsTC("ED_BANK3")
rsTA("AU_ACCOUNT3") = rsTC("ED_ACCOUNT3")
rsTA("AU_AMTDEPOSIT3") = rsTC("ED_AMTDEPOSIT3")
rsTA("AU_PCDEPOSIT3") = rsTC("ED_PCDEPOSIT3")
rsTA("AU_SUPCODE") = rsTC("ED_SUPCODE")
rsTA("AU_DDI") = rsTC("ED_DDI")
rsTA("AU_WCB") = rsTC("ED_WCB")
rsTA("AU_UNION") = rsTC("ED_UNION")
rsTA("AU_TD1") = rsTC("ED_TD1")
rsTA("AU_TD1DOL") = rsTC("ED_TD1DOL")
rsTA("AU_TD3") = rsTC("ED_TD3")
rsTA("AU_ProvAmt") = rsTC("ED_PROVAMT")
rsTA("AU_ExtraTax") = rsTC("ED_ExtraTax")
rsTA("AU_TD1CODE") = rsTC("ED_TD1CODE")
rsTA("AU_VACPC") = rsTC("ED_VACPC")
rsTA("AU_BUSNBR") = rsTC("ED_BUSNBR")
rsTA("AU_FDAY") = rsTC("ED_FDAY")
rsTA("AU_LDAY") = rsTC("ED_LDAY")
rsTA("AU_OMDAY") = rsTC("ED_OMERS")
rsTA("AU_INTEL") = rsTC("ED_INTEL")
Langs = Split(getLanguage(rsTC("ED_EMPNBR")), "|")
If Langs(0) <> "NoLang1" Then rsTA("AU_LANG1") = Langs(0) '0 is for ED_Lang1
If Langs(1) <> "NoLang2" Then rsTA("AU_LANG2") = Langs(1) '1 is for ED_Lang2
'George Apr 4,2006 #10574
rsTA("AU_EMAIL") = rsTC("ED_EMAIL")
rsTA("AU_WITHSPOUSE") = rsTC("ED_WITHSPOUSE")
rsTA("AU_EXPYEAR") = rsTC("ED_EXPYEAR")
rsTA("AU_ADMINBY") = rsTC("ED_ADMINBY")
rsTA("AU_CellPhone") = rsTC("ED_CellPhone")
rsTA("AU_PageNbr") = rsTC("ED_PageNbr")
rsTA("AU_SSN") = rsTC("ED_SSN")
rsTA("AU_DEPT_GL") = rsTC("ED_GLNO")
rsTA("AU_REGION") = rsTC("ED_REGION")
rsTA("AU_SECTION") = rsTC("ED_SECTION")
rsTA("AU_EMP") = rsTC("ED_EMP")
rsTA("AU_ORG") = rsTC("ED_ORG")
rsTA("AU_DRIVERLIC") = rsTC("ED_DRIVERLIC")
rsTA("AU_LICPLATE1") = rsTC("ED_LICPLATE1")
rsTA("AU_LICPLATE2") = rsTC("ED_LICPLATE2")
rsTA("AU_TYPEVEHICLE") = rsTC("ED_TYPEVEHICLE")
rsTA("AU_PARKPERMIT1") = rsTC("ED_PARKPERMIT1")
rsTA("AU_PARKPERMIT2") = rsTC("ED_PARKPERMIT2")
rsTA("AU_PAYROLL_ID") = rsTC("ED_PAYROLL_ID")
rsTA("AU_DEPTEDATE") = rsTC("ED_DEPTEDATE")
rsTA("AU_DIVEDATE") = rsTC("ED_DIVEDATE")
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA.Update

NewPayIDHREMPAudit = True

Exit Function

NewPayIDHREMPAudit_ERR:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js

End Function

Public Sub get_CopyToPayIDFormShort()
Dim xFormItem(5)
xFormItem(1) = "frmEPOSITION"
xFormItem(2) = "frmESALARY"

For X = 1 To NewHireForms.count: NewHireForms.Remove 1: Next
For X = 1 To 2
    NewHireForms.Add Trim(xFormItem(X))
Next
End Sub
