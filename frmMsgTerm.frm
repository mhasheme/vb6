VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmMsgTerm 
   Caption         =   "Termination Data"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   5805
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   1
      Top             =   2745
      Width           =   5805
      _Version        =   65536
      _ExtentX        =   10239
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
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
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
         Left            =   1200
         TabIndex        =   3
         Tag             =   "Save changes made"
         Top             =   0
         Visible         =   0   'False
         Width           =   735
      End
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
         Left            =   240
         TabIndex        =   2
         Tag             =   "Save changes made"
         Top             =   30
         Width           =   735
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5580
      MaxLength       =   25
      TabIndex        =   6
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5220
      MaxLength       =   25
      TabIndex        =   5
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   4950
      MaxLength       =   25
      TabIndex        =   4
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frmBasic 
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   -90
      TabIndex        =   0
      Top             =   -30
      Width           =   8235
      Begin VB.TextBox txtPublic 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2340
         MaxLength       =   9
         TabIndex        =   15
         Tag             =   "11-New Employee Number"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox txtEmpNum 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2340
         MaxLength       =   9
         TabIndex        =   12
         Tag             =   "11-New Employee Number"
         Top             =   1050
         Width           =   1185
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         Height          =   285
         Index           =   1
         Left            =   2040
         TabIndex        =   8
         Tag             =   "41-Termination Code "
         Top             =   630
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "TERM"
      End
      Begin INFOHR_Controls.DateLookup dlpTermDate 
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Tag             =   "41-Date Terminated"
         Top             =   270
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpJob 
         DataField       =   "JH_JOB"
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Tag             =   "01-Position code"
         Top             =   2160
         Visible         =   0   'False
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   25
         LookupType      =   5
      End
      Begin VB.Label lblPosTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Position Code"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   270
         TabIndex        =   18
         Top             =   2205
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblNote1 
         Caption         =   "lblNote1"
         Height          =   375
         Left            =   270
         TabIndex        =   17
         Top             =   1800
         Visible         =   0   'False
         Width           =   4575
      End
      Begin VB.Label lblPublic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Public"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   16
         Top             =   1470
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label lblEEName 
         Caption         =   "Message for duplicate "
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3810
         TabIndex        =   14
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblEmpNum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NEW Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   300
         TabIndex        =   13
         Top             =   1080
         Width           =   2340
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Reason"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   660
         Width           =   1710
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Termination Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Tag             =   "41-Date Terminated"
         Top             =   300
         Width           =   1830
      End
   End
End
Attribute VB_Name = "frmMsgTerm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mvarPenTermFlag As String
Dim mvarCmdCancel As Boolean

Private Sub cmdCancel_Click()
'Ticket #22409 Frank 08/08/2012
If glbWFC And mvarPenTermFlag = "WFC_SomkerChgReason" Then
    glbChgTermReason = txtPublic.Text
    glbSpouseSIN = "N"
    Unload Me
End If
End Sub

Private Sub cmdOK_Click()
Dim Msg As String, a%
Dim xlocDays As Integer
Dim xMsg As String

    'Ticket #22261 Franks 08/02/2012 - begin
    If glbSamuel And mvarPenTermFlag = "SamDateSection" Then
        If Len(dlpTermDate.Text) < 1 Then
            MsgBox (lStr("First Day") & "  is a required field")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(dlpTermDate.Text) Then
            MsgBox (lStr("First Day") & " is not a valid date.")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        glbChgTermDate = dlpTermDate.Text
        Unload Me
        Exit Sub
    End If
    If glbSamuel And mvarPenTermFlag = "SamDateRegion" Then
        If Len(dlpTermDate.Text) < 1 Then
            MsgBox (lStr("Last Day") & "  is a required field")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(dlpTermDate.Text) Then
            MsgBox (lStr("Last Day") & " is not a valid date.")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        glbChgTermDate = dlpTermDate.Text
        Unload Me
        Exit Sub
    End If
    'Ticket #22261 Franks 08/02/2012 - end

    'Ticket #22411 Franks 08/07/2012
    If glbWFC And mvarPenTermFlag = "WFCCOB_Change" Then
        If Len(dlpTermDate.Text) < 1 Then
            MsgBox "Benefit Effective Date is a required field"
            dlpTermDate.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(dlpTermDate.Text) Then
            MsgBox "Benefit Effective Date is not a valid date."
            dlpTermDate.SetFocus
            Exit Sub
        End If
        
        xlocDays = DateDiff("d", CVDate(dlpTermDate.Text), CVDate(Date))
        If xlocDays > 30 Then
            '"   The date is mandatory and must be within 30 days of "today". Use the same error message if the date is outside of the 30 day range.
            xMsg = "The Benefit Effective must be within 30 days of today"
            MsgBox xMsg
            dlpTermDate.SetFocus
            Exit Sub
        End If

        glbChgTermDate = dlpTermDate.Text
        Unload Me
        Exit Sub
    End If
    
    'Ticket #22409 Frank 08/08/2012
    If glbWFC And mvarPenTermFlag = "WFC_SomkerChgReason" Then
        If Len(txtPublic.Text) = 0 Then
            xMsg = "Please enter Reason."
            MsgBox xMsg
            txtPublic.SetFocus
            Exit Sub
        End If
        glbChgTermReason = txtPublic.Text
        glbSpouseSIN = "Y"
        Unload Me
        Exit Sub
    End If
    
    'Ticket #16395
    If glbWFC And mvarPenTermFlag = "Y" Then
        If Len(dlpTermDate.Text) < 1 Then
            MsgBox ("Pension Exit Date is a required field")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(dlpTermDate.Text) Then
            MsgBox ("Pension Exit Date is not a valid date.")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        glbChgTermDate = dlpTermDate.Text
        Unload Me
        Exit Sub
    End If

    'Ticket #16395
    If glbWFC And mvarPenTermFlag = "SpouseSIN" Then
        'glbChgTermDate = txtPublic.Text
        glbSpouseSIN = txtPublic.Text
        If Len(txtPublic.Text) < 1 Then
            Msg = "Are You Sure You Want To Leave Spouse SIN Blank? "
            
            a% = MsgBox(Msg, 36, "Confirm")
            If a% <> 6 Then
                txtPublic.SetFocus
                Exit Sub
            End If

        End If
        If Len(txtPublic.Text) > 0 Then
            If txtPublic.Text <> "999999999" Then
                If Not SIN_chk(txtPublic.Text) Then
                    MsgBox "Invalid SIN" & IIf(glbLinamar, "", "- if Unassigned set to 999-999-999")
                    txtPublic.SetFocus
                    Exit Sub
                End If
            End If
        End If
        Unload Me
        Exit Sub
    End If
    
    If glbWFC And mvarPenTermFlag = "WFCPosEnd_Change" Then 'Ticket #29183 Franks 09/13/2016
        If Len(clpJob.Text) = 0 Then
            MsgBox ("Position Code is a required field")
            clpJob.SetFocus
            Exit Sub
        End If
        If clpJob.Caption = "Unassigned" Then
            MsgBox ("Position Code is not a valid field")
            clpJob.SetFocus
            Exit Sub
        End If
        If Len(dlpTermDate.Text) < 1 Then
            MsgBox ("End Date  is a required field")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(dlpTermDate.Text) Then
            MsgBox "End Date is not a valid date."
            dlpTermDate.SetFocus
            Exit Sub
        End If
        If glbChgTermReason = clpJob.Text Then
            MsgBox "Position Code can not be same as the new Position Code(" & glbChgTermReason & ")"
            clpJob.SetFocus
            Exit Sub
        End If
        glbChgTermReason = clpJob.Text
        glbChgTermDate = dlpTermDate.Text
        Unload Me
        Exit Sub

    End If
    
    'Ticket #19266 Frank 10/27/2010
    If glbWFC And (mvarPenTermFlag = "NGS_EffectiveDate" Or mvarPenTermFlag = "NGS_TermDate") Then
        If mvarPenTermFlag = "NGS_EffectiveDate" Then
            If Len(dlpTermDate.Text) < 1 Then
                MsgBox (lStr("Other Date 1") & " is a required field")
                dlpTermDate.SetFocus
                Exit Sub
            End If
            If Not IsDate(dlpTermDate.Text) Then
                MsgBox (lStr("Other Date 1") & " is not a valid date.")
                dlpTermDate.SetFocus
                Exit Sub
            End If
        End If
        If mvarPenTermFlag = "NGS_TermDate" Then
            If Len(dlpTermDate.Text) < 1 Then
                MsgBox (lStr("Other Date 2") & " is a required field")
                dlpTermDate.SetFocus
                Exit Sub
            End If
            If Not IsDate(dlpTermDate.Text) Then
                MsgBox (lStr("Other Date 2") & " is not a valid date.")
                dlpTermDate.SetFocus
                Exit Sub
            End If
        End If
        glbChgTermDate = dlpTermDate
        'Unload Me
        GoTo end_line
    End If
    

    If glbWFC And glbEESection = "GREN" Then 'Get Payroll ID only
        If Len(txtEmpNum) <> 6 Then
            MsgBox ("Invalid format for Payroll ID. Format must be ######")
            txtEmpNum.SetFocus
            Exit Sub
        End If
        If lblEEName.Visible Then
            MsgBox ("Duplicate Payroll ID. ")
            txtEmpNum.SetFocus
            Exit Sub
        End If
        glbChgNewEmpnbr = txtEmpNum
    Else
        If Len(dlpTermDate.Text) < 1 Then
            MsgBox ("Termination Date is a required field")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        
        If Not IsDate(dlpTermDate.Text) Then
            MsgBox ("Termination Date is not a valid date.")
            dlpTermDate.SetFocus
            Exit Sub
        End If
        If Not glbMediPay Then    'Not MediPay
    
            If Len(clpCode(1).Text) < 1 Then
                MsgBox ("Termination Reason is a required field")
                clpCode(1).SetFocus
                Exit Sub
            End If
            If clpCode(1).Caption = "Unassigned" Then
                MsgBox ("Termination Reason is not a valid field")
                clpCode(1).SetFocus
                Exit Sub
            End If
            glbChgTermReason = clpCode(1)
            glbChgNewEmpnbr = txtEmpNum
        End If
        glbChgTermDate = dlpTermDate
        If glbCompSerial = "S/N - 2370W" Then
            If Len(txtEmpNum) = 0 Then
                MsgBox ("Payroll ID is a required field")
                txtEmpNum.SetFocus
                Exit Sub
            End If
            If lblEEName.Visible Then
                MsgBox ("Duplicate Payroll ID. ")
                txtEmpNum.SetFocus
                Exit Sub
            End If
            glbChgNewEmpnbr = txtEmpNum
        End If
    End If
end_line:
    Unload Me
End Sub

Private Sub Form_Activate()
If glbWFC And mvarPenTermFlag = "SpouseSIN" Then
    txtPublic.SetFocus
End If
End Sub

Private Sub Form_Load()
MDIMain.panHelp(0).Caption = "info:HR Message"
Me.Width = 6045
Me.Height = 3075
Call INI_Controls(Me)
'txtEmpNum = glbLEE_ID
If (glbWFC And glbEESection = "GREN") Or glbCompSerial = "S/N - 2370W" Then
    txtEmpNum = GetPayID(glbLEE_ID)
    lblTitle(0).Visible = False
    lblTitle(1).Visible = False
    dlpTermDate.Visible = False
    clpCode(1).Visible = False
    lblEmpNum.Caption = "Payroll ID"
ElseIf glbCompSerial = "S/N - 2370W" Then
    'txtEmpNum = GetPayID(glbLEE_ID)
    lblEmpNum.Caption = "Payroll ID"
ElseIf glbCompSerial = "S/N - 2382W" Or glbCompSerial = "S/N - 2380W" Then 'Namasco or VitalAire
    lblEmpNum.Visible = False
    txtEmpNum.Visible = False
    txtEmpNum = glbLEE_ID
Else
    txtEmpNum = glbLEE_ID
    If glbSoroc Or glbWFC Or glbVadim Then
        lblEmpNum.Visible = False
        lblEEName.Visible = False
        txtEmpNum.Visible = False
    End If
End If
If glbMediPay Then
    txtEmpNum = glbLEE_ID
    clpCode(1).Visible = False
    lblTitle(1).Visible = False
    lblEmpNum.Visible = False
    lblEEName.Visible = False
    txtEmpNum.Visible = False
End If

'Ticket #16395
If glbWFC And mvarPenTermFlag = "Y" Then
    lblTitle(0).Caption = "Pension Exit Date"
    frmMsgTerm.Caption = "Pension Exit Date"
    clpCode(1).Visible = False
    lblTitle(1).Visible = False
End If

'Ticket #22411 Franks 08/07/2012
If glbWFC And mvarPenTermFlag = "WFCCOB_Change" Then
    lblTitle(0).Caption = "Benefit Effective Date"
    frmMsgTerm.Caption = "Enter Benefit Effective Date"
    clpCode(1).Visible = False
    lblTitle(1).Visible = False
End If

'Ticket #22409 Frank 08/08/2012
If glbWFC And mvarPenTermFlag = "WFC_SomkerChgReason" Then
    frmMsgTerm.Caption = "Confirm"
    lblNote1.Top = 300: lblNote1.Height = 750: lblNote1.Visible = True
    lblTitle(0).Top = 1080: 'dlpTermDate.Top = 1080
    lblPublic = "Reason"
    lblPublic.Visible = True
    txtPublic.Visible = True
    txtPublic.MaxLength = 200 '50 'Ticket #23768 Franks 05/14/2013
    txtPublic.Left = 1200
    txtPublic.Width = 4500 ' 3500 'Ticket #23768 Franks 05/14/2013
    
    cmdOK.Caption = "Yes"
    cmdCancel.Caption = "No"
    cmdCancel.Visible = True
    txtPublic.TabIndex = 0
    
    lblTitle(0).Visible = False
    lblTitle(1).Visible = False
    dlpTermDate.Visible = False
    clpCode(1).Visible = False
End If

'Ticket #22261 Franks 07/27/2012
If glbSamuel And mvarPenTermFlag = "SamDateSection" Then
    lblTitle(0).Caption = lStr("First Day")
    frmMsgTerm.Caption = "Enter Change Date"
    clpCode(1).Visible = False
    lblTitle(1).Visible = False
End If
'Ticket #22261 Franks 08/02/2012
If glbSamuel And mvarPenTermFlag = "SamDateRegion" Then
    lblTitle(0).Caption = lStr("Last Day")
    frmMsgTerm.Caption = "Enter Change Date"
    clpCode(1).Visible = False
    lblTitle(1).Visible = False
End If

'Ticket #16395 Frank 12/17/2009 - WFC Pension Outstanding Tasks By Dec1009.doc
If glbWFC And mvarPenTermFlag = "SpouseSIN" Then
    lblPublic.Caption = "Spouse SIN"
    lblPublic.Top = lblTitle(1).Top
    txtPublic.Top = clpCode(1).Top
    lblPublic.Visible = True
    txtPublic.Visible = True
    frmMsgTerm.Caption = "Spouse SIN"
    lblTitle(0).Visible = False
    dlpTermDate.Visible = False
    clpCode(1).Visible = False
    lblTitle(1).Visible = False
    lblEEName.Visible = False
    txtEmpNum.Visible = False
End If

'Ticket #19266
If glbWFC And (mvarPenTermFlag = "NGS_EffectiveDate" Or mvarPenTermFlag = "NGS_TermDate") Then
    If mvarPenTermFlag = "NGS_EffectiveDate" Then
        frmMsgTerm.Caption = lStr("Other Date 1") '"NGS Effective Date"
        lblTitle(0).Caption = lStr("Other Date 1") '"NGS Effective Date"
    End If
    If mvarPenTermFlag = "NGS_TermDate" Then
        frmMsgTerm.Caption = lStr("Other Date 2") '"NGS Term Date"
        lblTitle(0).Caption = lStr("Other Date 2") '"NGS Term Date"
    End If
    dlpTermDate.Text = glbChgTermDate
    lblTitle(0).Top = lblTitle(1).Top
    dlpTermDate.Top = clpCode(1).Top
    lblTitle(0).Visible = True
    dlpTermDate.Visible = True
    lblPublic.Visible = False 'True
    txtPublic.Visible = False 'True
    clpCode(1).Visible = False
    lblTitle(1).Visible = False
    lblEEName.Visible = False
    txtEmpNum.Visible = False
End If

If glbWFC And mvarPenTermFlag = "WFCPosEnd_Change" Then 'Ticket #29183 Franks 09/13/2016
    lblPosTitle.Top = 500
    clpJob.Top = 500
    clpJob.TextBoxWidth = 1315
    lblPosTitle.Visible = True
    clpJob.Visible = True
    lblTitle(0).Top = 1080
    dlpTermDate.Top = 1080
    lblTitle(1).Visible = False 'Termination Reason
    clpCode(1).Visible = False
    
    frmMsgTerm.Caption = "Select the Old Position"
    lblTitle(0).Caption = "End Date"
    
End If

If glbWFC And mvarPenTermFlag = "WFC_SomkerChgReason" Then
Else
    cmdCancel.Visible = mvarCmdCancel
End If

If glbCompSerial = "S/N - 2460W" Then 'Oshawa Public Libraries = Ticket #25323 Franks 12/16/2014
    lblEmpNum.Visible = False
    txtEmpNum.Visible = False
End If

End Sub

Private Function GetPayID(xEnpNo)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ, xStr
    xStr = ""
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        xStr = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
    GetPayID = xStr
End Function

Private Sub txtEmpNum_Change()


Dim rsEmp As New ADODB.Recordset
lblEEName() = ""
lblEEName().Visible = False
If Not IsNumeric(txtEmpNum) Then Exit Sub
If glbWFC And glbEESection = "GREN" Then
        SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_SECTION='GREN' AND ED_PAYROLL_ID='" & txtEmpNum & "' "
        SQLQ = SQLQ & " AND ED_EMPNBR <> " & glbLEE_ID
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsEmp.EOF Then
            lblEEName = "This Payroll ID already exists"
            lblEEName.Visible = True
        End If
ElseIf glbCompSerial = "S/N - 2370W" Then
        SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_PAYROLL_ID='" & txtEmpNum & "' "
        'SQLQ = SQLQ & " AND ED_EMPNBR <> " & glbLEE_ID
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsEmp.EOF Then
            lblEEName = "This Payroll ID already exists"
            lblEEName.Visible = True
        End If
Else
    If Len(txtEmpNum) > 0 And txtEmpNum <> glbLEE_ID Then
        rsEmp.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & txtEmpNum, gdbAdoIhr001, adOpenForwardOnly
        If Not rsEmp.EOF Then
            lblEEName = "This number already exists"
            lblEEName.Visible = True
        End If
    End If
End If
End Sub
Private Sub txtEmpNum_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Public Property Get PenTermDate() As String
    PenTermDate = mvarPenTermFlag
End Property

Public Property Let PenTermDate(ByVal vData As String)
    mvarPenTermFlag = vData
End Property
Public Property Get cmdCancelShow() As Boolean
    cmdCancelShow = mvarCmdCancel
End Property
Public Property Let cmdCancelShow(ByVal vData As Boolean)
    mvarCmdCancel = vData
End Property

