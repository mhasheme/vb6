VERSION 5.00
Begin VB.Form frmEHSF7WhatsMissing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "What is Missing in Form 7?"
   ClientHeight    =   2085
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox comWhatsMissing 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmMsgDialog1.frx":0000
      Left            =   120
      List            =   "frmMsgDialog1.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   7335
   End
   Begin VB.CommandButton cmdClose 
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
      Left            =   3173
      TabIndex        =   0
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblForm7 
      Caption         =   "What's Missing in WSIB Form 7:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmEHSF7WhatsMissing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Generate What's Missing List
    comWhatsMissing.Clear
    Call WSIBForm7_Whats_Missing
End Sub

Private Sub WSIBForm7_Whats_Missing()
    Dim rsEMP As New ADODB.Recordset
    Dim rsCompMst As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim rsSal As New ADODB.Recordset
    Dim rsHS As New ADODB.Recordset
    Dim rsStatCat As New ADODB.Recordset
    Dim rsForm7Sec As New ADODB.Recordset
    Dim SQLQ  As String
    Dim xMissingFlg As Boolean
    
    xMissingFlg = False
    
    SQLQ = "SELECT ED_EMPNBR, ED_FNAME, ED_SURNAME, ED_ADDR1, ED_CITY, ED_PROV, ED_PCODE, ED_ORG,"
    SQLQ = SQLQ & " ED_DOB, ED_PHONE, ED_DOH, ED_SEX, ED_SIN, ED_BUSNBR, ED_EMP, ED_PT, ED_TD1DOL,"
    SQLQ = SQLQ & " ED_PROVAMT"
    SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEMP.EOF Then
        comWhatsMissing.AddItem "Employee Demographics"
        comWhatsMissing.AddItem "Status/Dates Information"
        comWhatsMissing.AddItem "Banking Information"
        xMissingFlg = True
    Else
        xMissingFlg = False
    End If
    
    SQLQ = "SELECT JH_EMPNBR, JH_JOB, JH_SDATE, JH_DHRS FROM HR_JOB_HISTORY"
    SQLQ = SQLQ & " WHERE JH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND JH_CURRENT <> 0"
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsJOB.EOF Then
        comWhatsMissing.AddItem "Employee Position Information"
        xMissingFlg = True
    Else
        xMissingFlg = False
    End If

    SQLQ = "SELECT SH_EMPNBR, SH_EDATE, SH_SALARY, SH_SALCD, SH_WHRS FROM HR_SALARY_HISTORY"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND SH_CURRENT <> 0"
    rsSal.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsSal.EOF Then
        comWhatsMissing.AddItem "Employee Salary Information"
        xMissingFlg = True
    Else
        xMissingFlg = False
    End If
    
    SQLQ = "SELECT * FROM HR_OCC_HEALTH_SAFETY"
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    'SQLQ = SQLQ & " AND EC_CASE =" & Data1.Recordset!EC_CASE
    SQLQ = SQLQ & " AND EC_CASE =" & glbF7CaseNo
    rsHS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsHS.EOF Then
        comWhatsMissing.AddItem "Employee Health & Safety Information"
        xMissingFlg = True
    Else
        xMissingFlg = False
    End If
    
    SQLQ = "SELECT * FROM HR_OHS_FORM7_SECTIONS"
    SQLQ = SQLQ & " WHERE F7_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND F7_CASE =" & glbF7CaseNo
    rsForm7Sec.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsForm7Sec.EOF Then
        comWhatsMissing.AddItem "Additional Form 7 Sections"
        xMissingFlg = True
    Else
        xMissingFlg = False
    End If
    
    If glbF7FirmAcct <> "" And glbF7FirmAcctNo <> "" Then
        SQLQ = "SELECT * FROM HR_OHS_COMPANY_MASTER"
        'SQLQ = SQLQ & " WHERE EY_FIRM_ACCT = '" & Data1.Recordset!EC_FIRM_ACCT & "'"
        'SQLQ = SQLQ & " AND EY_FIRM_ACCT_NUM ='" & Data1.Recordset!EC_FIRM_ACCT_NUM & "'"
        SQLQ = SQLQ & " WHERE EY_FIRM_ACCT = '" & glbF7FirmAcct & "'"
        SQLQ = SQLQ & " AND EY_FIRM_ACCT_NUM ='" & glbF7FirmAcctNo & "'"
        rsCompMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsCompMst.EOF Then
            comWhatsMissing.AddItem "Employer Information, i.e. Firm / Account #"
            xMissingFlg = True
        Else
            xMissingFlg = False
        End If
    End If
    
    If Not rsEMP.EOF Then
        SQLQ = "SELECT SC_WORKER_TYPE FROM HR_EMPLOYEE_MATRIX"
        SQLQ = SQLQ & " WHERE SC_EMP = '" & rsEMP("ED_EMP") & "'"
        SQLQ = SQLQ & " AND SC_PT = '" & rsEMP("ED_PT") & "'"
        rsStatCat.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsStatCat.EOF Then
            comWhatsMissing.AddItem "Form 7 Employee Type Matrix"
            xMissingFlg = True
        Else
            xMissingFlg = False
        End If
    End If
    
    If xMissingFlg Then GoTo Close_All
    
    
    'Whats Missing:
    If Not rsEMP.EOF Then
        'A:
        If IsNull(rsEMP("ED_SIN")) Then comWhatsMissing.AddItem "Section A: Social Insurance Number"
        If IsNull(rsEMP("ED_SURNAME")) Then comWhatsMissing.AddItem "Section A: Last Name"
        If IsNull(rsEMP("ED_FNAME")) Then comWhatsMissing.AddItem "Section A: First Name"
        If IsNull(rsEMP("ED_ADDR1")) Then comWhatsMissing.AddItem "Section A: Address "
        If IsNull(rsEMP("ED_CITY")) Then comWhatsMissing.AddItem "Section A: City/Town"
        If IsNull(rsEMP("ED_PROV")) Then comWhatsMissing.AddItem "Section A: Province"
        If IsNull(rsEMP("ED_PCODE")) Then comWhatsMissing.AddItem "Section A: Postal Code"
        If IsNull(rsHS("EC_JBSDATE")) Then comWhatsMissing.AddItem "Section A: Length of time in this Position - Position Start Date"
    End If

    If glbF7FirmAcct <> "" And glbF7FirmAcctNo <> "" Then
        If Not rsCompMst.EOF Then
            'B.
            If IsNull(rsCompMst("EY_TRADLEGAL_NAME")) Then comWhatsMissing.AddItem "Section B: Trade and Legal Name"
            If IsNull(rsCompMst("EY_FIRM_ACCT_NUM")) Then comWhatsMissing.AddItem "Section B: Firm or Account Number"
            If IsNull(rsCompMst("EY_MAIL_ADDRESS")) Then comWhatsMissing.AddItem "Section B: Mailing Address"
            If IsNull(rsCompMst("EY_CITY")) Then comWhatsMissing.AddItem "Section B: City/Town"
            If IsNull(rsCompMst("EY_PROV")) Then comWhatsMissing.AddItem "Section B: Province"
            If IsNull(rsCompMst("EY_PCODE")) Then comWhatsMissing.AddItem "Section B: Postal Code"
            If IsNull(rsCompMst("EY_PHONE")) Then comWhatsMissing.AddItem "Section B: Telephone"
            If IsNull(rsCompMst("EY_RATE_GRP_NUM")) Then comWhatsMissing.AddItem "Section B: Rate Group Number"
            If IsNull(rsCompMst("EY_CLASS_UNIT_CODE")) Then comWhatsMissing.AddItem "Section B: Classification Unit Code"
            If IsNull(rsCompMst("EY_BUSINESS_DESC")) Then comWhatsMissing.AddItem "Section B: Description of Business Activity"
        End If
    Else
        comWhatsMissing.AddItem "Section B: Employer Information"
    End If
    
    If Not rsHS.EOF Then
        'C:
        If IsNull(rsHS("EC_OCCDATE")) Then comWhatsMissing.AddItem "Section C: 1. Date of accident/Awareness of illness"
        If IsNull(rsHS("EC_OCCTM")) Then comWhatsMissing.AddItem "Section C: 1. Hour of accident/Awareness of illness"
        If IsNull(rsHS("EC_DATENOT")) Then comWhatsMissing.AddItem "Section C: 1. Date reported to employer"
        If IsNull(rsHS("EC_TIMNOT")) Then comWhatsMissing.AddItem "Section C: 1. Hour reported to employer"
        If IsNull(rsHS("EC_EMPNOT")) Then comWhatsMissing.AddItem "Section C: 2. Who was the accident/illness reported to?"
        If IsNull(GetEmpData(rsHS("EC_EMPNOT"), "ED_BUSNBR")) Then comWhatsMissing.AddItem "Section C: 2. Who was the accident/illness reported to Telephone?"
        If IsNull(rsHS("EC_CLASS")) Then comWhatsMissing.AddItem "Section C: 3. Was the accident/illness"
        If IsNull(rsHS("EC_TYPE")) Then comWhatsMissing.AddItem "Section C: 4. Type of accident/illness:"
        If IsNull(rsHS("EC_PREMISES")) Then comWhatsMissing.AddItem "Section C: 7. Did the accident/illness happen on the employer's premises?"
        If IsNull(rsHS("EC_OUTSIDE_PROV")) Then comWhatsMissing.AddItem "Section C: 8. Did the accident/illness happen outside the Province of Ontario?"
        If IsNull(rsHS("EC_WITNESS")) Then comWhatsMissing.AddItem "Section C: 9. Are you aware of any witnesses or other employees involved in this accident/illness?"
        If IsNull(rsHS("EC_INDIV_RESP")) Then comWhatsMissing.AddItem "Section C: 10. Was any individual, who does not work for your firm, partially or totally responsible for this accident/illness?"
        If IsNull(rsHS("EC_SIMILAR_INJ")) Then comWhatsMissing.AddItem "Section C: 11. Are you aware of any prior similar or related problem, injury or condition?"
        
        'D:
        If IsNull(rsHS("EC_PHYS1_VISIT")) Then comWhatsMissing.AddItem "Section D: 1. Did the worker receive health care for this injury?"
        If IsNull(rsHS("EC_PHYS1_NOTIFIED")) Then comWhatsMissing.AddItem "Section D: 2. When did the employer learn that the worker received health care?"
        If IsNull(rsHS("EC_FAPROVIDED")) Then comWhatsMissing.AddItem "Section D: 3. Where was the worker treated for this injury?"
        If IsNull(rsHS("EC_PHYSNM")) Then comWhatsMissing.AddItem "Section D: 3. Name, address and phone number of health professional..."
    End If
    
    If Not rsForm7Sec.EOF Then
        'E
        If IsNull(rsForm7Sec("F7_RETURNED_TO")) Then comWhatsMissing.AddItem "Section E: 1. After the day of accident/awareness of illness, this worker:"
        If IsNull(rsForm7Sec("F7_CONFIRM_BY")) Then comWhatsMissing.AddItem "Section E: 2. This Lost Time-No Lost Time-Modified Worker information was confirmed by:"
    End If

    If Not rsForm7Sec.EOF Then
        'F
        'Only required if Section E. - '....Modified Work' selected
        If rsForm7Sec("F7_RETURNED_TO") = "M" Then
            If IsNull(rsForm7Sec("F7_LIMITATION")) Then comWhatsMissing.AddItem "Section F: 1. Have you been provided with work limitations for this workers injury?"
            If IsNull(rsForm7Sec("F7_DISCUSSED")) Then comWhatsMissing.AddItem "Section F: 2. Has modified work been discussed with this worker?"
            If IsNull(rsForm7Sec("F7_OFFERED")) Then comWhatsMissing.AddItem "Section F: 3. Has modified work been offered to this worker?"
            If IsNull(rsForm7Sec("F7_ACCEPT_DECLINE")) Then comWhatsMissing.AddItem "Section F: 4. Who is responsible for arranging worker's return to work"
        End If
    End If
    
    If Not rsEMP.EOF Then
        If rsStatCat.EOF Then
            'G:
            If IsNull(rsStatCat("SC_WORKER_TYPE")) Then comWhatsMissing.AddItem "Section G: 1. Is this worker..."
        End If
    End If
    
    'Only required if Section E. - 'Has lost time and/or earnings' is selected
'    If Not rsForm7Sec.EOF Then
'        If rsForm7Sec("F7_RETURNED_TO") = "L" Then
'            If Not rsEMP.EOF Then
'                'H:
'                If IsNull(rsEMP("ED_TD1DOL")) Then comWhatsMissing.AddItem "Section H: 1. Net Claim Code or Amount - Federal"
'                If IsNull(rsEMP("ED_PROVAMT")) Then comWhatsMissing.AddItem "Section H: 1. Net Claim Code or Amount - Provincial"
'            End If
'        End If
'    End If
    
    If Not rsForm7Sec.EOF Then
        'H
        'Only required if Section E. - 'Has lost time and/or earnings' is selected
        If rsForm7Sec("F7_RETURNED_TO") = "L" Then
            If IsNull(rsForm7Sec("F7_FED_AMT")) Then comWhatsMissing.AddItem "Section H: 1. Net Claim Code or Amount - Federal"
            If IsNull(rsForm7Sec("F7_PROV_AMT")) Then comWhatsMissing.AddItem "Section H: 1. Net Claim Code or Amount - Provincial"
            If IsNull(rsForm7Sec("F7_VAC_PAY")) Then comWhatsMissing.AddItem "Section H: 2. Vacation pay - on each cheque?"
            'If Not IsNull(rsForm7Sec("F7_VACPC")) Then  'Provide percentage %
            If IsNull(rsForm7Sec("F7_LAST_WORK_DATE")) Then comWhatsMissing.AddItem "Section H: 3. Date and hour last worked"
            'If Not IsNull(rsForm7Sec("F7_LAST_WORK_TIME")) Then comWhatsMissing.AddItem "Section H: 3. Hour"
            If IsNull(rsForm7Sec("F7_LAST_DAY_WORK_FTIME")) Then comWhatsMissing.AddItem "Section H: 4. Normal working hours on last day worked From"
            If IsNull(rsForm7Sec("F7_LAST_DAY_WORK_TTIME")) Then comWhatsMissing.AddItem "Section H: 4. Normal working hours on last day worked To"
            If IsNull(rsForm7Sec("F7_LAST_DAY_ACT_EARN")) Then comWhatsMissing.AddItem "Section H: 5. Actual earnings for last day worked $"
            If IsNull(rsForm7Sec("F7_LAST_DAY_NORM_EARN")) Then comWhatsMissing.AddItem "Section H: 6. Normal earnings for last day worked $"
            If IsNull(rsForm7Sec("FY_WORKER_PAID")) Then comWhatsMissing.AddItem "Section H: 7. Advances on wages: Is the worker being paid while he/she recovers?"
                'If Not IsNull(rsForm7Sec("F7_WORKER_FTREGOTHR")) Then   'If Yes
                'If rsForm7Sec("F7_WORKER_FTREGOTHR") = "O" Then     'If Other then Name
                    'If IsNull(rsForm7Sec("F7_WORKER_OTHER")) Then
        End If
    End If

    If Not rsForm7Sec.EOF Then
        'I
        If rsForm7Sec("F7_RETURNED_TO") = "L" Then
            If IsNull(rsForm7Sec("F7_WORKSCH")) Then comWhatsMissing.AddItem "Section I: (Complete either A, B or C. Do not include overtime shifts.)"
        End If
    End If
        
    
    If comWhatsMissing.ListCount > 0 Then
        comWhatsMissing.ListIndex = 0
    End If

Close_All:
    rsStatCat.Close
    Set rsStatCat = Nothing
    
    If glbF7FirmAcct <> "" And glbF7FirmAcctNo <> "" Then
        rsCompMst.Close
        Set rsCompMst = Nothing
    End If
    
'    rsForm7Sec.Close
'    Set rsForm7Sec = Nothing
    
    rsHS.Close
    Set rsHS = Nothing
    
    rsSal.Close
    Set rsSal = Nothing
    
    rsJOB.Close
    Set rsJOB = Nothing
    
    rsEMP.Close
    Set rsEMP = Nothing

End Sub

