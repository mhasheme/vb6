Attribute VB_Name = "Intergradation"

Option Explicit
Public Enum PassStatus
    NONE = 0
    Demographices = 1
    Status = 2
    Position = 3
    Salary = 4
    Banking = 5
    Contacts = 6
    Termination = 7
    Rehire = 8
    Attendance = 9
    Benefit = 10
    Accrual = 11
    HoulyEntitlement = 12
    SalaryGirdMaster = 13
    PositionMaster = 14
End Enum
Public Enum FieldTypeEnum
    aString
    aDate
    aNumber
End Enum

Sub Passing_Bank_Changes(Banks As Collection, xEMPNBR, Optional xPayID)
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub
If glbVadim Then
    Call Passing_Bank_Vadim(Banks, xEMPNBR, xPayID)
End If
End Sub

Sub Passing_Attendance_Changes(HRAtts As Collection, UptType, xEMPNBR, Optional xPayID)
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub
If glbVadim Then
   'NOT IN USE IN VB CODE. HAS TRIGGER IN DATABASE
   ' Call Passing_Attendance_Vadim(HRAtts, UptType, xEmpnbr, xPayID)
End If
End Sub

Sub Passing_Changes(HRChanges As Collection, PStatus As PassStatus, UptType, UptDate, Optional xEMPNBR, Optional xPayID)
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub
If glbVadim Then
    Call Passing_Changes_Vadim(HRChanges, PStatus, UptType, UptDate, xEMPNBR, xPayID)
End If
End Sub

Sub AddNewPayrollEmp(PStatus As PassStatus, UptDate, xEMPNBR, xPayID)
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub
If glbVadim Then
    Call AddNewVadimEmp(PStatus, UptDate, xEMPNBR, xPayID)
End If
End Sub

Sub TermPayrollEmp(UptDate, xEMPNBR, Optional xPayID, Optional PStatus As PassStatus)
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub
If glbVadim Then
    Call TermVadimEmp(UptDate, xEMPNBR, xPayID, PStatus)
End If
End Sub

Sub DeletePayrollEmp(UptDate, xEMPNBR, Optional xPayID)
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub
If glbVadim Then
    Call DeleteVadimEmp(UptDate, xEMPNBR, xPayID)
End If
End Sub

Function isChanged_Field(HRChanges As Collection, oldValue, InField As Object, Optional NumberComp As Boolean) As Boolean
Dim NewValue
Dim HRField
Dim rsJOB As New ADODB.Recordset
Dim isDiff As Boolean
Dim xHRChange As HRChange

Set xHRChange = New HRChange

isChanged_Field = False

If glbVadim Then
    If TypeOf InField Is FieldInfo Then
        HRField = InField.fdName
        NewValue = InField.fdValue
    ElseIf InField.name = "txtLambtonJob" Then
        HRField = "JH_JOB"
        NewValue = InField
    Else
        If TypeOf InField Is ADODB.Field Then
            HRField = InField.name
        Else
            HRField = InField.DataField
        End If
        NewValue = InField
        If HRField = "ED_PAYROLL_ID" Then HRField = "JH_PAYROLL_ID"
        If HRField = "ED_DEPTNO" Then HRField = "JH_DEPTNO"
        If HRField = "ED_GLNO" Then HRField = "JH_GLNO"
        If HRField = "ED_ORG" Then HRField = "JH_ORG"
        If HRField = "ED_EMP" Then HRField = "JH_EMP"
        
        'City of Kawartha Lakes
        If glbCompSerial = "S/N - 2363W" Then
            If HRField = "SH_JOB" Then HRField = "JH_JOB"
            If HRField = "ED_REGION" Then HRField = "JH_PAYROLL_CATEGORY"
            If HRField = "ED_LDAY" Then HRField = "TERM:RETI-DATE"  'Ticket #18734
        ElseIf glbCompSerial = "S/N - 2375W" Then  'City of Timmins
            If HRField = "SH_JOB" Then HRField = "JH_JOB"
            If HRField = "ED_NORMALR" Then HRField = "TERM:RETI-DATE"
        End If
        'City of Niagara Falls
        If glbCompSerial = "S/N - 2276W" Then
            If HRField = "ED_VADIM2" Then HRField = "ED_VADIM2_SICK"
            If HRField = "ED_CHDSUP" Then HRField = "SICK_ACCR_HOURS"
            'Ticket #20053
            'If HRField = "ED_ADMINBY" Then HRField = "EMP_CLASS_CODE"
            If HRField = "ED_BENEFIT_GROUP" Then HRField = "EMP_CLASS_CODE"
            
            If HRField = "ED_GARN" Then HRField = "SICK_ACCR_MAXIMUM"
            If HRField = "ED_PAYFREQ" Then HRField = "ED_VACPC_1"
        End If
        'Town of Aurora
        If glbCompSerial = "S/N - 2378W" Then
            'Ticket #20931 - as per mapping documentation
            If HRField = "ED_BENEFIT_GROUP" Then HRField = "EMP_CLASS_CODE"
            'If HRField = "ED_ADDR2" Then HRField = "UNIT_NUM"
            If HRField = "ED_DIV" Then HRField = "JH_PAYROLL_CATEGORY"
            If HRField = "ED_GLNO" Then HRField = "JH_GLNO"
            If HRField = "ED_LDAY" Then HRField = "TERM:LO-DATE"
        End If
        'Ticket #19113 - District Municipality of Muskoka
        If glbCompSerial = "S/N - 2373W" Then
            'If HRField = "JH_ORG" Then HRField = "JH_PAYROLL_CATEGORY" 'Since it updates two fields in Vadim taken case from Pass_Changes_Vadim
            'If HRField = "ED_MSTAT" Then HRField = "EMP_CLASS_CODE"
        End If
        'Ticket #23795 - Town of Lasalle
        If glbCompSerial = "S/N - 2379W" Then
            If HRField = "ED_CELLPHONE" Then HRField = "ED_BUSNBR"
            If HRField = "ED_DIV" Then HRField = "JH_PAYROLL_CATEGORY"
            If HRField = "ED_LDAY" Then HRField = "TERM:LO-DATE"
            If HRField = "ED_PENSION" Then HRField = "ED_OMERS_1"
        End If
        'Ticket #24375 - Town of Greater Napanee
        If glbCompSerial = "S/N - 2447W" Then
            'If HRField = "ED_EMPTYPE" Then HRField = "EMP_CLASS_CODE"  'They are not using this field in Vadim.
            If HRField = "ED_LDAY" Then HRField = "TERM:LO-DATE"
            If HRField = "ED_LTHIRE" Then HRField = "TERM:RETI-DATE"
            If HRField = "ED_PT" Then HRField = "JH_PAYROLL_CATEGORY"
            'If HRField = "ED_ADMINBY" Then HRField = "EMP_DEFAULT_JOB" 'They are not using this field in Vadim.
        End If
        'Ticket #24996 - City of Campbell River
        If glbCompSerial = "S/N - 2458W" Then
            If HRField = "ED_CELLPHONE" Then HRField = "ED_BUSNBR"  'Ticket #25469
            If HRField = "ED_SECTION" Then HRField = "EMP_CLASS_CODE"
            If HRField = "ED_BENEFIT_GROUP" Then HRField = "EMP_DEFAULT_JOB"
            If HRField = "ED_PT" Then HRField = "JH_PAYROLL_CATEGORY"
            If HRField = "ED_LDAY" Then HRField = "TERM:LO-DATE"
            'Ticket #28990 - They don't want this to be transferred from info:HR anymore
            'If HRField = "ED_USER_NUM1" Then HRField = "SICK_ACCR_HOURS"
            'If HRField = "ED_USER_NUM2" Then HRField = "VAC_ACC_HOURS"
            If HRField = "ED_ORGT1" Then HRField = "JB_DESCR"
        End If
    End If
Else
    If TypeOf InField Is FieldInfo Then
        HRField = InField.fdName
        NewValue = InField.fdValue
    ElseIf TypeOf InField Is ADODB.Field Then
        HRField = InField.name
        NewValue = InField
    Else
        HRField = InField.DataField
        NewValue = InField
    End If
End If

If NumberComp Then
    If IsNull(oldValue) Then oldValue = 0
    If IsNull(NewValue) Then NewValue = 0
    isDiff = Val(oldValue) <> Val(NewValue)
Else
    If IsNull(oldValue) Then oldValue = ""
    If IsNull(NewValue) Then NewValue = ""
    isDiff = oldValue <> NewValue
End If
If isDiff Then
    If glbVadim Then
'        If HRField = "SH_EDATE" Or HRField = "SH_NEXTDAT" Then
'            If oldValue = "01/01/01" Then oldValue = ""
'            If oldValue = NewValue Then Exit Function
'        End If
        xHRChange.HRField = HRField
        xHRChange.NewValue = NewValue
        xHRChange.oldValue = oldValue
        HRChanges.Add xHRChange, HRField
    End If
    isChanged_Field = True
End If
End Function

Function isChanged_Salary(HRSalary As Collection, oldValue, InField As Object, Optional NumberComp As Boolean) As Boolean
Dim isDiff As Boolean
Dim NewValue
Dim HRField
Dim xHRChange As HRChange
Set xHRChange = New HRChange
isChanged_Salary = False


If TypeOf InField Is FieldInfo Then
    HRField = InField.fdName
    NewValue = InField.fdValue
ElseIf TypeOf InField Is ADODB.Field Then
    HRField = InField.name
    NewValue = InField
Else
    HRField = InField.DataField
    NewValue = InField
    
    'Ticket #26891 - Commenting it because Passing_Salary_Vadim is now looking for SH_TOTAL in the HRSalary collection
    'If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
    '    If HRField = "SH_TOTAL" Then HRField = "SH_SALARY"
    'End If
End If
If NumberComp Then
    If IsNull(oldValue) Then oldValue = 0
    If Not IsNumeric(NewValue) Then NewValue = 0
    isDiff = Val(oldValue) <> Val(Format(NewValue, "@"))
Else
    If IsNull(oldValue) Then oldValue = ""
    isDiff = oldValue <> Format(NewValue, "@")
End If

If isDiff Then isChanged_Salary = True
xHRChange.HRField = HRField
xHRChange.NewValue = NewValue
xHRChange.oldValue = oldValue
HRSalary.Add xHRChange, HRField

End Function

Function Add_DFLT_Values(HRChanges As Collection, HRField, DefaultValue) As Boolean

Dim NewValue, oldValue
Dim xHRChange As New HRChange
Dim X
For X = 1 To HRChanges.count
    If UCase(HRChanges(X).HRField) = UCase(HRField) Then
        Exit Function
    End If
Next

oldValue = Null
NewValue = DefaultValue

xHRChange.HRField = HRField
xHRChange.NewValue = NewValue
xHRChange.oldValue = oldValue
HRChanges.Add xHRChange, HRField

End Function
Function isChanged_Attendance(HRAtts As Collection, oldValue, InField As Object, Optional NumberComp As Boolean) As Boolean
'NOT IN USE IN VB CODE. HAS TRIGGER IN DATABASE
'Dim isDiff As Boolean
'Dim NewValue
'Dim HRField
'Dim xHRChange As New HRChange
'isChanged_Attendance = False
'
'NewValue = InField
'If TypeOf InField Is ADODB.Field Then
'    HRField = InField.name
'Else
'    HRField = InField.DataField
'End If
'If NumberComp Then
'    If IsNull(OldValue) Then OldValue = 0
'    If Not IsNumeric(NewValue) Then NewValue = 0
'    isDiff = Val(OldValue) <> Val(Format(NewValue, "@"))
'Else
'    If IsNull(OldValue) Then OldValue = ""
'    isDiff = OldValue <> Format(NewValue, "@")
'End If
'
'If isDiff Then isChanged_Attendance = True
'xHRChange.HRField = HRField
'xHRChange.NewValue = NewValue
'xHRChange.OldValue = OldValue
'HRAtts.Add xHRChange, HRField

End Function

Function isChanged_Bank(Banks As Collection, oldValue, InField As Object, Optional NumberComp As Boolean) As Boolean
Dim isDiff As Boolean
Dim NewValue

isChanged_Bank = False

NewValue = InField

If NumberComp Then
    If IsNull(oldValue) Then oldValue = 0
    isDiff = Val(oldValue) <> Val(Format(NewValue, "@"))
Else
    If IsNull(oldValue) Then oldValue = ""
    isDiff = oldValue <> Format(NewValue, "@")
End If

If isDiff Then isChanged_Bank = True
If glbVadim Then Call Add_Bank_Info(Banks, oldValue, InField)
End Function

Sub Add_Bank_Info(Banks As Collection, oldValue, InField As Object)
Dim NewValue
Dim OldField, HRField
Dim xBNo
Dim X
Dim xHRChange As New HRChange

NewValue = InField
If TypeOf InField Is ADODB.Field Then
    HRField = InField.name
Else
    HRField = InField.DataField
End If

HRField = Mid(HRField, 4)

xBNo = Val(Right(HRField, 1))
If xBNo = 0 Then
    xBNo = 1
Else
    HRField = Left(HRField, Len(HRField) - 1)
End If
If Banks.count <> xBNo Then
    Dim BankInfo As New Collection
    Banks.Add BankInfo
    For X = 1 To BankInfo.count: BankInfo.Remove 1: Next
End If

If (HRField = "PCDEPOSIT") Then 'Or (HRField = "ED_PCDEPOSIT2") Or (HRField = "ED_PCDEPOSIT3") Then
    If NewValue <> "" Then
        NewValue = NewValue / 100
    End If
    If oldValue <> "" Then
        oldValue = oldValue / 100
    End If
End If

xHRChange.HRField = HRField
xHRChange.NewValue = NewValue
xHRChange.oldValue = oldValue
Banks(xBNo).Add xHRChange, HRField

End Sub


Function getPayrollIDs(xEMPNBR, Optional xPayID, Optional FromEmp As Boolean) As String
Dim X
Dim rsJOB As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset

getPayrollIDs = ""
If IsMissing(xPayID) Then xPayID = "" Else xPayID = Format(xPayID, "@")
If xPayID <> "" Then
    getPayrollIDs = xPayID
Else
    If FromEmp Then
        rsEmp.Open "SELECT ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR=" & xEMPNBR, gdbAdoIhr001, adOpenForwardOnly
        If Not rsEmp.EOF Then
           xPayID = Format(rsEmp("ED_PAYROLL_ID"), "@")
           getPayrollIDs = xPayID
        End If
        rsEmp.Close
    Else
        rsJOB.Open "SELECT JH_PAYROLL_ID FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEMPNBR, gdbAdoIhr001, adOpenForwardOnly
        If rsJOB.EOF Then
            rsEmp.Open "SELECT ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR=" & xEMPNBR, gdbAdoIhr001, adOpenForwardOnly
            If Not rsEmp.EOF Then
               xPayID = Format(rsEmp("ED_PAYROLL_ID"), "@")
               getPayrollIDs = xPayID
            End If
            rsEmp.Close
        Else
            getPayrollIDs = ""
            Do Until rsJOB.EOF
                xPayID = Format(rsJOB("JH_PAYROLL_ID"), "@")
                getPayrollIDs = getPayrollIDs & "|" & xPayID
                rsJOB.MoveNext
            Loop
            getPayrollIDs = Mid(getPayrollIDs, 2)
        End If
        rsJOB.Close
    End If
End If
    
End Function
Function getGLNo(xEMPNBR, Optional xJob) As String
Dim X
Dim rsJOB As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset

getGLNo = ""
If IsMissing(xJob) Then xJob = "" Else xJob = Format(xJob, "@")
If xJob = "" Then
    rsEmp.Open "SELECT ED_GLNO FROM HREMP WHERE ED_EMPNBR=" & xEMPNBR, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
       getGLNo = Format(rsEmp("ED_GLNO"), "@")
    End If
    rsEmp.Close
Else
    rsJOB.Open "SELECT JH_GLNO FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEMPNBR & " AND JH_JOB='" & xJob & "'", gdbAdoIhr001, adOpenForwardOnly
    If Not rsJOB.EOF Then
       getGLNo = Format(rsJOB("JH_GLNO"), "@")
    End If
    rsJOB.Close
End If
    
End Function

Function Employee_Mass_Update_Integration(oldCode, newCode, xPlantCode, xFunction)
Dim CommandStr
If glbAdv Then
    'xFunction =
    'Employee Dept Mass Update
    'Employee JobCode Mass Update
    'Employee Category Mass Update
    
    'Ticket #25911 Franks 03/17/2015 - Position Codes are not in AT, it is Job Code
    'WFC doesn't need this in Ver 8.1
    If glbWFC And xFunction = "Employee JobCode Mass Update" Then
        Exit Function
    End If
    
    CommandStr = "Advanced Tracker"
    CommandStr = CommandStr & "," & xFunction 'Employee JobCode Mass Update"
    If Len(xPlantCode) > 0 Then
        CommandStr = CommandStr & "/" & xPlantCode
    Else
        CommandStr = CommandStr & "/ALL"
    End If
    CommandStr = CommandStr & "," & oldCode
    CommandStr = CommandStr & ","
    CommandStr = CommandStr & newCode
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
End Function
Function Employee_PositionDel_Integration(updEMPID, jobcodeID, startDate, Optional DeleteRecord, Optional TermSEQ)
Dim CommandStr
If glbAdv Then
    CommandStr = "Advanced Tracker"
    CommandStr = CommandStr & ",Employee Position Delete"
    If Len(glbPlantCode) > 0 Then
        CommandStr = CommandStr & "/" & glbPlantCode
    Else
        CommandStr = CommandStr & "/ALL"
    End If
    CommandStr = CommandStr & "," & updEMPID
    'for newkey
    CommandStr = CommandStr & ","
    CommandStr = CommandStr & jobcodeID
    CommandStr = CommandStr & ","
    CommandStr = CommandStr & startDate
    If Not IsMissing(DeleteRecord) Then
        CommandStr = CommandStr & "," & DeleteRecord
        If glbWFC Then
            Call Pause(5)
        End If
    Else
        CommandStr = CommandStr & ","
    End If
    If Not IsMissing(TermSEQ) Then CommandStr = CommandStr & TermSEQ
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
End Function

Function Employee_GL_Dist_Integration(updEMPID, Optional newEMPID, Optional DeleteRecord, Optional TermSEQ, Optional GLCode)
Dim CommandStr
Dim xEmpPlantCode As String
Dim xTerSEQ As Integer

If glbGP Then 'Ticket #24518 Franks 04/21/2014
    CommandStr = "Great Plains"
    CommandStr = CommandStr & ",Employee GL Dist"
    If Not IsMissing(GLCode) Then
        CommandStr = CommandStr & "/" & GLCode
    End If
    CommandStr = CommandStr & "," & updEMPID
    
    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newEMPID) Then CommandStr = CommandStr & newEMPID
    'If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
    
    CommandStr = CommandStr & ","   'Hemu
    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & DeleteRecord
    CommandStr = CommandStr & ","   'Hemu
    If Not IsMissing(TermSEQ) Then CommandStr = CommandStr & TermSEQ
    CommandStr = CommandStr & ","
    CommandStr = CommandStr & ","
    CommandStr = CommandStr & gsGPHold
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If

End Function

Function Employee_Master_Integration(updEMPID, Optional newEMPID, Optional DeleteRecord, Optional TermSEQ, Optional UptType)
Dim CommandStr
Dim xEmpPlantCode As String
Dim xTerSEQ As Integer

Call CGLUpdate(updEMPID, newEMPID, DeleteRecord) ' will only do it for CGL Manufacturing

If glbAdv Then
    CommandStr = "Advanced Tracker"
    CommandStr = CommandStr & ",Employee Master"
    If Len(glbPlantCode) > 0 Then
        ''Ticket #17903
        'xEmpPlantCode = getEmpPlantCode(updEMPID)
        'If Len(xEmpPlantCode) > 0 Then
        '    If xEmpPlantCode = glbPlantCode Then
        '        CommandStr = CommandStr & "/" & glbPlantCode
        '    Else
        '        CommandStr = CommandStr & "/" & xEmpPlantCode
        '    End If
        'Else
        '    CommandStr = CommandStr & "/" & glbPlantCode
        'End If
        CommandStr = CommandStr & "/" & glbPlantCode
        If glbWFC Then 'Ticket #24337 Franks 10/01/2013
            'Ticket #27609 Franks 10/07/2015 - If this employee is Not in Tracker then do not call AT Integration
            If IsWFCNotInTracker(updEMPID) Then
                Exit Function
            End If
            If Not IsWFCAdvConnected Then
                MsgBox "Unable to connect to the Tracker Database." & Chr(10) & "Please notify system administrator and try again later"
                Exit Function
            End If
        End If
    Else
        CommandStr = CommandStr & "/ALL"
    End If
    CommandStr = CommandStr & "," & updEMPID
    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newEMPID) Then CommandStr = CommandStr & newEMPID
    If Not IsMissing(DeleteRecord) Then
        CommandStr = CommandStr & "," & DeleteRecord
        If glbWFC Then
            Call Pause(5)
        End If
    Else
        CommandStr = CommandStr & ","
    End If
    If Not IsMissing(TermSEQ) Then CommandStr = CommandStr & TermSEQ
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
    
If glbGP Then
    CommandStr = "Great Plains"
    CommandStr = CommandStr & ",Employee Master"
    If Not IsMissing(UptType) Then
        CommandStr = CommandStr & "/" & UptType
    End If
    CommandStr = CommandStr & "," & updEMPID
    
    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newEMPID) Then CommandStr = CommandStr & newEMPID
    'If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
    
    CommandStr = CommandStr & ","   'Hemu
    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & DeleteRecord
    CommandStr = CommandStr & ","   'Hemu
    If Not IsMissing(TermSEQ) Then CommandStr = CommandStr & TermSEQ
    CommandStr = CommandStr & ","
    CommandStr = CommandStr & ","
    CommandStr = CommandStr & gsGPHold
    If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #22342 Franks 10 / 23 / 2012
        If Not IsMissing(TermSEQ) Then xTerSEQ = TermSEQ Else xTerSEQ = 0
        'get Div code from employee, not the login user, so the change by any user can be sent to GP
        glbPlantCode = getEmpDivCode(updEMPID, xTerSEQ)
        If Len(glbPlantCode) > 0 Then
            CommandStr = CommandStr & "," & glbPlantCode
        Else
            CommandStr = CommandStr & ",ALL"
        End If
    End If
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
    
If glbMediPay Then
    'Employee must be manually terminated in MediPay
    'So do not integrate terminated employees Ticket #14752
    If Not IsMissing(TermSEQ) Then
        Exit Function
    End If
    
    If glbCompSerial = "S/N - 2242W" Then 'CCAC London #9014
        CommandStr = "MediPay|glbBatchNumber"
    Else
        CommandStr = "MediPay"
    End If
    
    CommandStr = CommandStr & ",Employee Master"
    CommandStr = CommandStr & "," & updEMPID
    
    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newEMPID) Then
        CommandStr = CommandStr & "," & newEMPID
    End If
    
    If Not IsMissing(DeleteRecord) Then
        CommandStr = CommandStr & "," & DeleteRecord
    End If
    
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
    
End Function

Function getEmpPlantCode(xEmpID)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    SQLQ = "SELECT ED_EMPNBR, ED_SECTION FROM HREMP WHERE ED_EMPNBR = " & xEmpID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_SECTION")) Then
            retVal = rsEmp("ED_SECTION")
        End If
    End If
    rsEmp.Close
    getEmpPlantCode = retVal
End Function

Function Employee_Transfered_MediPay(updEMPID, Optional newEMPID, Optional DeleteRecord)
Dim CommandStr

    CommandStr = "MediPay"
    CommandStr = CommandStr & ",Employee Transfered"
    CommandStr = CommandStr & "," & updEMPID
    
    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newEMPID) Then
        CommandStr = CommandStr & "," & newEMPID
    End If
    
    If Not IsMissing(DeleteRecord) Then
        CommandStr = CommandStr & "," & DeleteRecord
    End If
    
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus

    
End Function

Function Employee_GP_NewBenefitDeduction_Integration(updEMPID) 'Ticket #26654 Franks 02/10/2015
Dim CommandStr
If glbGP Then
    CommandStr = "Great Plains"
    CommandStr = CommandStr & ",NewBenDeduction"
    CommandStr = CommandStr & "," & updEMPID
    CommandStr = CommandStr & "," '& UserID
    CommandStr = CommandStr & "," 'delete
    CommandStr = CommandStr & "," '& UseWRK
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
End Function

Function Employee_GP_BenefitDeduction_Integration(updEMPID, UserID, UseWRK As Boolean)
'UseWRK - if use working table
Dim CommandStr
If glbGP Then
    CommandStr = "Great Plains"
    CommandStr = CommandStr & ",BenefitDeduction"
    CommandStr = CommandStr & "," & updEMPID
    CommandStr = CommandStr & "," & UserID
    CommandStr = CommandStr & "," 'delete
    CommandStr = CommandStr & "," & UseWRK
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If

End Function

Function Employee_Benefit_Integration(updEMPID, Optional newEMPID, Optional DeleteRecord, Optional NewRecord, Optional BEID)
Dim CommandStr
If glbMediPay Then
    CommandStr = "MediPay"
    CommandStr = CommandStr & ",Benefit Changes"
    CommandStr = CommandStr & "," & updEMPID
    
    'for newkey
    If Not IsMissing(newEMPID) Then
        CommandStr = CommandStr & "," & newEMPID
    Else
        CommandStr = CommandStr & "," & updEMPID
    End If
    
    If Not IsMissing(DeleteRecord) Then
        CommandStr = CommandStr & "," & DeleteRecord
    Else
        CommandStr = CommandStr & ",false"
    End If
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If

End Function

Function Employee_GP_OMERS_Integration(updEMPID, Optional newEMPID, Optional DeleteRecord, Optional NewRecord, Optional BEID)
Dim CommandStr
'comment by Frank 12/08/2009
'GP OMERS Calculation on Union Add, change on Status/Date screen.
'see Ticket #17643
'If glbGP Then
'    If glbCompSerial = "S/N - 2172W" Then 'County of Lanark
'        CommandStr = "Great Plains"
'        CommandStr = CommandStr & ",GP_OMERS"
'        CommandStr = CommandStr & "," & updEMPID
'
'        'for newkey
'        If Not IsMissing(newEMPID) Then
'            CommandStr = CommandStr & "," & newEMPID
'        Else
'            CommandStr = CommandStr & "," & updEMPID
'        End If
'
'        If Not IsMissing(DeleteRecord) Then
'            CommandStr = CommandStr & "," & DeleteRecord
'        Else
'            CommandStr = CommandStr & ",false"
'        End If
'        Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
'    End If
'End If

End Function


Function Position_Integration(updEMPID, Optional newEMPID, Optional DeleteRecord, Optional NewRecord, Optional SHID) 'George Mar 7,2006 #9965 Added 2 new optional paras "NewRecord","SHID"
Dim CommandStr

If glbMediPay Then
    CommandStr = "MediPay"
    CommandStr = CommandStr & ",Position Changes"
    CommandStr = CommandStr & "," & updEMPID
    
    'for newkey
    If Not IsMissing(newEMPID) Then
        CommandStr = CommandStr & "," & newEMPID
    Else
        CommandStr = CommandStr & "," & updEMPID
    End If
    
    If Not IsMissing(DeleteRecord) Then
        CommandStr = CommandStr & "," & DeleteRecord
    Else
        CommandStr = CommandStr & ",false"
    End If
'    CommandStr = CommandStr & "," & OldClientNum
'
'    If Not IsMissing(NewClientNum) Then
'        CommandStr = CommandStr & "," & NewClientNum
'    Else
'        CommandStr = CommandStr & "," & OldClientNum
'    End If
'
'    If Not IsMissing(InsertRecord) Then
'        CommandStr = CommandStr & "," & InsertRecord
'    Else
'        CommandStr = CommandStr & ",false"
'    End If
    
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
    
End Function

Function Salary_Integration(updEMPID, Optional newEMPID, Optional DeleteRecord, Optional NewRecord, Optional SHID) 'George Mar 7,2006 #9965 Added 2 new optional paras "NewRecord","SHID"
Dim CommandStr

If glbMediPay Then
    CommandStr = "MediPay"
    CommandStr = CommandStr & ",Salary Changes"
    CommandStr = CommandStr & "," & updEMPID
    
    'for newkey
    If Not IsMissing(newEMPID) Then
        CommandStr = CommandStr & "," & newEMPID
    Else
        CommandStr = CommandStr & "," & updEMPID
    End If
    
    If Not IsMissing(DeleteRecord) Then
        CommandStr = CommandStr & "," & DeleteRecord
    Else
        CommandStr = CommandStr & ",false"
    End If
'    CommandStr = CommandStr & "," & OldClientNum
'
'    If Not IsMissing(NewClientNum) Then
'        CommandStr = CommandStr & "," & NewClientNum
'    Else
'        CommandStr = CommandStr & "," & OldClientNum
'    End If
'
'    If Not IsMissing(InsertRecord) Then
'        CommandStr = CommandStr & "," & InsertRecord
'    Else
'        CommandStr = CommandStr & ",false"
'    End If
    
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If

'If glbAdv Then
'    CommandStr = "Advanced Tracker"
'    CommandStr = CommandStr & ",Employee Master"
'    CommandStr = CommandStr & "," & updEMPID
'    'for newkey
'    CommandStr = CommandStr & ","
'    If Not IsMissing(newEMPID) Then CommandStr = CommandStr & newEMPID
'
'    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
'    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
'End If
'
If glbGP Then
    CommandStr = "Great Plains"
    CommandStr = CommandStr & ",Salary Changes"
    CommandStr = CommandStr & "," & updEMPID

    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newEMPID) Then CommandStr = CommandStr & newEMPID

    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
    If Not IsMissing(NewRecord) Then CommandStr = CommandStr & "," & NewRecord 'George Mar 7,2006 #9965
    If Not IsMissing(SHID) Then CommandStr = CommandStr & "," & SHID 'George Mar 7,2006 #9965
    CommandStr = CommandStr & "," & gsGPHold
    If glbCompSerial = "S/N - 2443W" Then 'Walters Inc Ticket #22342 Franks 10/24/2012
        'If Not IsMissing(TermSEQ) Then xTerSEQ = TermSEQ Else xTerSEQ = 0
        'get Div code from employee, not the login user, so the change by any user can be sent to GP
        glbPlantCode = getEmpDivCode(updEMPID, 0)
        If Len(glbPlantCode) > 0 Then
            CommandStr = CommandStr & "," & glbPlantCode
        Else
            CommandStr = CommandStr & ",ALL"
        End If
    End If
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If
    
    
End Function

Function Bank_Integration(updEMPID, Optional newEMPID, Optional DeleteRecord, Optional TermSEQ)
Dim CommandStr
    If glbMediPay Then
        'Employee must be manually terminated in MediPay
        'So do not integrate terminated employees Ticket #14752
        If Not IsMissing(TermSEQ) Then
            Exit Function
        End If
        
        If glbCompSerial = "S/N - 2242W" Then 'CCAC London #9014
            CommandStr = "MediPay|glbBatchNumber"
        Else
            CommandStr = "MediPay"
        End If
        
        CommandStr = CommandStr & ",Bank Changes"
        CommandStr = CommandStr & "," & updEMPID
        
        'for newkey
        CommandStr = CommandStr & ","
        If Not IsMissing(newEMPID) Then
            CommandStr = CommandStr & "," & newEMPID
        End If
        
        If Not IsMissing(DeleteRecord) Then
            CommandStr = CommandStr & "," & DeleteRecord
        End If
        
        Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
    End If

End Function

Function Attendance_Master_Integration(updKey, Optional newKey, Optional DeleteRecord, Optional ATNonPaidHours)
Dim CommandStr

Dim FuncitonName
If glbAdv Then
    'Ticket #24268 Franks 10/02/2013 - begin
    'For WFC and Mitchell Plastics, they use WFC_Attend_To_AT function
    If glbWFC Then Exit Function
    If glbMitchellPlastics Then Exit Function
    'Ticket #24268 Franks 10/02/2013 - end
    
    CommandStr = "Advanced Tracker"
    If Not IsMissing(ATNonPaidHours) Then
        CommandStr = CommandStr & ",Export Attendance_NonPaid"
    Else
        CommandStr = CommandStr & ",Export Attendance"
    End If
    If Len(glbPlantCode) > 0 Then
        CommandStr = CommandStr & "/" & glbPlantCode
    Else
        CommandStr = CommandStr & "/ALL"
    End If
    CommandStr = CommandStr & "," & updKey
    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
    
    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If

'If glbMediPay Then
'    CommandStr = "MediPay"
'    CommandStr = CommandStr & ",Export Attendance"
'    CommandStr = CommandStr & "," & updKey
'    'for newkey
'    CommandStr = CommandStr & ","
'    If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
'
'    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
'    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
'End If
End Function
Function Entitlements_Master_Integration(updKey, Optional newKey, Optional DeleteRecord)
Dim CommandStr

Dim FuncitonName
If glbAdv Then
    CommandStr = "Advanced Tracker"
    CommandStr = CommandStr & ",Entitlements Update"
    If Len(glbPlantCode) > 0 Then
        CommandStr = CommandStr & "/" & glbPlantCode
    Else
        CommandStr = CommandStr & "/ALL"
    End If
    CommandStr = CommandStr & "," & updKey
    'for newkey
    CommandStr = CommandStr & ","
    If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
    
    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
End If

End Function
Function Codes_Master_Integration(CodeName, updKey, Optional newKey, Optional DeleteRecord)
Dim CommandStr

Dim FuncitonName
    Select Case CodeName
    Case "POSITION"
        FuncitonName = "Position Master"
    Case "DEPT"
        FuncitonName = "Department Master"
    Case "ADRE"
        FuncitonName = "Attendance Codes Master"
    Case "EDPT"
        FuncitonName = "Category Master"
    Case "HOLIDAY"
        FuncitonName = "Holiday Master"
    End Select
    If glbWFCFullRights And FuncitonName = "Holiday Master" Then
        CommandStr = "Advanced Tracker"
        CommandStr = CommandStr & "," & FuncitonName

        If Len(newKey) > 0 Then
            CommandStr = CommandStr & "/" & newKey
        Else
            CommandStr = CommandStr & "/ALL"
        End If
        CommandStr = CommandStr & "," & updKey
        'for newkey
        CommandStr = CommandStr & ","
        'If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
        
        If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
        Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
    End If
    If glbAdv And Len(FuncitonName) > 0 Then 'Or glbWFCFullRights Then 'MZ doesn't want the super user to pass AT Master data 'Ticket #11964
        CommandStr = "Advanced Tracker"
        CommandStr = CommandStr & "," & FuncitonName
        If FuncitonName = "Holiday Master" Then
            If Len(newKey) > 0 Then
                CommandStr = CommandStr & "/" & newKey
            Else
                CommandStr = CommandStr & "/ALL"
            End If
            CommandStr = CommandStr & "," & updKey
            'for newkey
            CommandStr = CommandStr & ","
            'If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
            
            If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
        Else
            If Len(glbPlantCode) > 0 Then
                CommandStr = CommandStr & "/" & glbPlantCode
            Else
                CommandStr = CommandStr & "/ALL"
            End If
            CommandStr = CommandStr & "," & updKey
            'for newkey
            CommandStr = CommandStr & ","
            If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
            
            If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
        End If
        Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
    End If
        
    If glbGP Then
        CommandStr = "Great Plains"
        CommandStr = CommandStr & "," & FuncitonName
        CommandStr = CommandStr & "," & updKey
        'for newkey
        CommandStr = CommandStr & ","
        If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
        
        If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
        Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
    End If
    
    If glbMediPay Then
        CommandStr = "MediPay"
        CommandStr = CommandStr & "," & FuncitonName
        CommandStr = CommandStr & "," & updKey
        'for newkey
        CommandStr = CommandStr & ","
        If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
        
        If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
        Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
    End If

'Comment by Frank Apr 7, 2005, WFC donesn't want to Intergrade with Bonus System,
'We will use export files to Bonus System
'Else 'For WFC
'    Select Case CodeName
'    Case "POSITION"
'        FuncitonName = "Position Master"
'    Case "EDLC"
'        FuncitonName = "Location"
'    Case "EDAB"
'        FuncitonName = "HR Manager"
'    End Select
'
'    CommandStr = "Bonus System"
'    CommandStr = CommandStr & "," & FuncitonName
'    CommandStr = CommandStr & "," & updKey
'    'for newkey
'    CommandStr = CommandStr & ","
'    If Not IsMissing(newKey) Then CommandStr = CommandStr & newKey
'
'    If Not IsMissing(DeleteRecord) Then CommandStr = CommandStr & "," & DeleteRecord
'    Shell glbIntegrationEXE & " " & CommandStr, vbNormalFocus
'End If

End Function

Function isATIncluded(xEMPNBR)
isATIncluded = True
If glbLambton Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & xEMPNBR
    SQLQ = SQLQ & " AND ED_LOC IN (SELECT PARA_VALUE FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='Advanced Tracker' AND PARA_CATEGORY2='Integration Selection' AND PARA_NAME='Location')"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsEmp.EOF Then isATIncluded = False
End If
End Function

Function getEmpDivCode(xEmpNo, xTermSEQ)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    If xTermSEQ = 0 Then
        SQLQ = "SELECT ED_EMPNBR, ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_DIV,TERM_SEQ FROM Term_HREMP WHERE TERM_SEQ = " & xTermSEQ & " "
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_DIV")) Then
            retVal = rsEmp("ED_DIV")
        End If
    End If
    getEmpDivCode = retVal
End Function

Function isTransferAT(Product_Info, xFunction, xMultiPayCode) 'Ticket #24124 Franks 07/24/2013
Dim rsSetup As New ADODB.Recordset
Dim SQLQ
isTransferAT = False
SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & Product_Info & "' AND PARA_CATEGORY2='Integration Setup' AND PARA_NAME='" & xFunction & "' "
If Product_Info = "Advanced Tracker" Then
    If Len(xMultiPayCode) > 0 Then
        If xMultiPayCode = "ALL" Then
        Else
            SQLQ = SQLQ & "AND PARA_MULTIPAY_CODE = '" & xMultiPayCode & "' "
        End If
    End If
End If
rsSetup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If rsSetup.EOF Then Exit Function
isTransferAT = rsSetup("PARA_VALUE") = "1"
End Function

'Ticket #24124 Franks 07/24/2013
Public Function OtherDatabaseInte(Product_Info As String, xMultiPayCode, Optional Emp_No) As String
Dim SQLQ
Dim xVersion
Dim xDatabaseName
Dim xDatabaseServer
Dim xUsername
Dim xPassword
Dim xDatabasePath
Dim rsDataSetup As New ADODB.Recordset
On Error GoTo err_OtherDatabaseInte

OtherDatabaseInte = ""

SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & Product_Info & "' AND PARA_CATEGORY2='Database Setup' "
'Get the Payroll Code for multi Database setup, such as WFC
If Product_Info = "Advanced Tracker" Then
    If xMultiPayCode = "ALL" Then
        'SQLQ = SQLQ & "AND PARA_MULTIPAY_CODE IS NULL "
    Else
        If Len(xMultiPayCode) > 0 Then
        SQLQ = SQLQ & "AND PARA_MULTIPAY_CODE = '" & xMultiPayCode & "' "
        End If
    End If
End If
If Product_Info = "Great Plains" Then
    If glbCompSerial = "S/N - 2443W" Then ' SERIAL_WaltersInc Then 'Ticket #22342
        SQLQ = SQLQ & "AND PARA_MULTIPAY_CODE = '" & xMultiPayCode & "' "
    End If
End If
rsDataSetup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
Do Until rsDataSetup.EOF
    If rsDataSetup("PARA_NAME") = "Version_Info" Then
       xVersion = rsDataSetup("PARA_VALUE")
    Else
       If xVersion = "MS SQL Server" Then
            If rsDataSetup("PARA_NAME") = "Database_Name" Then
                xDatabaseName = rsDataSetup("PARA_VALUE")
            End If
            If rsDataSetup("PARA_NAME") = "Database_Server" Then
                xDatabaseServer = rsDataSetup("PARA_VALUE")
            End If
            If rsDataSetup("PARA_NAME") = "User_Name" Then
                xUsername = rsDataSetup("PARA_VALUE")
            End If
            If rsDataSetup("PARA_NAME") = "Password" Then
                xPassword = rsDataSetup("PARA_VALUE")
            End If
            
        Else
            If rsDataSetup("PARA_NAME") = "Database_Path" Then
                xDatabasePath = rsDataSetup("PARA_VALUE")
                xDatabasePath = xDatabasePath & IIf(Right(xDatabasePath, 1) = "\", "", "\")
            End If
            If rsDataSetup("PARA_NAME") = "Database_Name" Then
                xDatabaseName = rsDataSetup("PARA_VALUE")
            End If
        End If
    End If
    rsDataSetup.MoveNext
Loop
rsDataSetup.Close
If xVersion = "MS SQL Server" Then
    OtherDatabaseInte = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & xUsername & ";Password=" & xPassword & ";Initial Catalog=" & xDatabaseName & ";Data Source=" & xDatabaseServer
ElseIf xVersion = "MS Access" Then
    OtherDatabaseInte = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=petman;Data Source=" & xDatabasePath & xDatabaseName
End If
Exit Function
err_OtherDatabaseInte:
    OtherDatabaseInte = ""
End Function

'Ticket #24124 Franks 07/24/2013
Public Function TableNamePrefix(xMultiPayCode) As String
Dim SQLQ
Dim rsPrefix As New ADODB.Recordset
On Error GoTo err_TableNamePrefix
TableNamePrefix = ""

SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='Advanced Tracker' AND PARA_CATEGORY2='Database Setup' AND PARA_NAME='Table Name Prefix' "
If xMultiPayCode = "ALL" Then
    SQLQ = SQLQ & "AND PARA_MULTIPAY_CODE IS NULL "
Else
    If Len(xMultiPayCode) > 0 Then
    SQLQ = SQLQ & "AND PARA_MULTIPAY_CODE = '" & xMultiPayCode & "' "
    End If
End If
rsPrefix.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If Not rsPrefix.EOF Then
    TableNamePrefix = rsPrefix("PARA_VALUE") & ""
End If
rsPrefix.Close
Exit Function
err_TableNamePrefix:
    TableNamePrefix = ""
End Function

