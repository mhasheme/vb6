Attribute VB_Name = "Vadim_Intergradation"
Option Explicit
Global IVMap As New Collection
Global VDTables As New Collection
Global Vadim_PayType_field
Global Vadim_EmpType_field
Global Vadim_PayType_TABLName
Global Vadim_EmpType_TABLName



Public Enum VadimTableNum
    EMPLOYEE
    CLIENT
    PAY_DEPOSIT_DISTRIBUTION
End Enum

Public Type PayCodeInfoType
    PayCode As String
    PayType As String
    PayTypeID As String
    PayFreq As String
End Type


Sub VadimInterface(xBatchID, ByVal xPayID, HRField, ByVal OldValue, ByVal NewValue, Optional VDClause)
Dim rsFace As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim VDTableFields
Dim xMapField
Dim CvtHRField 'converted HR field
Dim X, Y
Dim FindMatchField As Boolean
On Error Resume Next

FindMatchField = False
VDTableFields = Split(IVMap(HRField), ",")
If Err.Number = 0 Then
    FindMatchField = True
Else
    For X = 1 To IVMap.count
        VDTableFields = Split(IVMap(X), ",")
        xMapField = Split(VDTableFields(0), ":")
        If xMapField(0) = "VIT" Then ' for default value
            If HRField = getVITField(xMapField(1)) Then
                FindMatchField = True
                Exit For
            End If
        End If
        If xMapField(0) = "RESET" Then
            If HRField = xMapField(1) Then
                FindMatchField = True
                Exit For
            End If
        End If
        'City of Niagara Falls, Town of Aurora
        If glbCompSerial = "S/N - 2276W" Or glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2458W" Then
            If xMapField(0) = "DFLT" Then
                If HRField = xMapField(1) Then
                    FindMatchField = True
                    Exit For
                End If
            End If
        End If
    Next

End If

On Error GoTo VadimInterface_err
    If Not FindMatchField Then Exit Sub
    If HRField = "JH_PAYROLL_ID" And IsNull(NewValue) Then
        NewValue = xPayID
    End If
    For Y = 1 To UBound(VDTableFields)
        If VDTables(VDTableFields(Y)).fdType = aDate Then
            If IsDate(OldValue) Then
                OldValue = Format(OldValue, "YYYY/MM/DD")
            Else
                OldValue = Null
            End If
            If IsDate(NewValue) Then
                NewValue = Format(NewValue, "YYYY/MM/DD")
            Else
                'City of Timmins - When Null or "Null" is passed to remove the Termination Date - Vadim is not accepting it
                If glbCompSerial = "S/N - 2375W" And HRField = "TERM:EMP_DATE" Then
                    NewValue = ""
                Else
                    NewValue = Null
                End If
            End If
        End If
        Call PassDataToVadim(xBatchID, VDTableFields(Y), xPayID, OldValue, NewValue, VDClause)
    Next
Exit Sub
VadimInterface_err:
MsgBox Err.Description
Resume Next
End Sub

Sub Passing_Changes_Vadim(HRChanges As Collection, PStatus As PassStatus, UptType, UptDate, xEmpnbr, Optional xPayID)
Dim X, Y
Dim PayIDs
Dim HRField
Dim OldValue, NewValue
Dim xBatchID
Dim xPayType
Dim xCompNo, xOCompNo
Dim xDept
Dim xDeptName
Dim xUnion
Dim xNxtGrdFlg As Boolean
Dim xTmpValue

On Error GoTo Passing_Bank_Changes_Vadim_Err
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(PStatus) Then Exit Sub

PayIDs = Split(getPayrollIDs(xEmpnbr, xPayID), "|")
For X = 0 To UBound(PayIDs)
    If glbCompSerial = "S/N - 2276W" And (PStatus = "3" Or PStatus = "4") Then   'City of Niagara Falls - Ticket #14285
        xBatchID = AddBatchVadim(UptType, UptDate) 'add interface header
    ElseIf (glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2447W") And PStatus = "4" Then    'Ticket #24565,Ticket #25412 - District Municipality of Muskoka, Town of Greater Napanee - Future dated Salary changes transfer as of Future Date
        xBatchID = AddBatchVadim(UptType, UptDate) 'add interface header
    Else
        xBatchID = AddBatchVadim(UptType) 'add interface header
    End If
    
    'start adding information to the interface detail table
    For Y = 1 To HRChanges.count
        HRField = HRChanges(Y).HRField
        OldValue = HRChanges(Y).OldValue
        NewValue = HRChanges(Y).NewValue
        
        If HRField = "ED_WCB" Or HRField = "ED_WCBCODE" Or UCase(HRField) = "ED_PROVAMT" Then
            'City of Kawartha Lakes
            If glbCompSerial = "S/N - 2363W" Then
                If HRField = "ED_WCB" Or UCase(HRField) = "ED_PROVAMT" Then
                    Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
                End If
            Else
                Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
            End If
        Else
            'Dist. of Muskoka
            If (glbCompSerial = "S/N - 2373W") Then
                If Right(HRField, 4) = "_ORG" Then
                    Call VadimInterface(xBatchID, PayIDs(X), "JH_PAYROLL_CATEGORY", OldValue, NewValue)
                
                    'WO# 129311 - CUPE Code - Create CUPE Union Due in iCity for Union Codes 1181, 1810 and PPT
                    If NewValue = "1810" Or NewValue = "1811" Or NewValue = "PPT" Or _
                       OldValue = "1810" Or OldValue = "1811" Or OldValue = "PPT" Then
                       Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
                    End If
                End If
            End If
        
            'City of Kawartha Lakes
            If glbCompSerial = "S/N - 2363W" And HRField = "ED_VADIM1" Then
                 Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
            
            'City of Timmins or City of Niagara Falls
            ElseIf (glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W") And HRField = "ED_VADIM2" And PStatus = Banking Then
                Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
        
            'City of Niagara Falls or District Municipality of Muskoka (Ticket #19113)
            'Ticket #25412 - Town of Greater Napanee
            'Get the Code to pass from Code Matrix
            ElseIf (glbCompSerial = "S/N - 2276W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2447W") And Right(HRField, 4) = "_ORG" Then
                If HRField = "ED_ORG" Or HRField = "JH_ORG" Then
                    NewValue = CodeMatrix("EDOR", CStr(NewValue), "")
                    OldValue = CodeMatrix("EDOR", CStr(OldValue), "")
                End If
                'If NewValue = "" Then NewValue = "Null"
                Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
            
            'Town of Lasalle
            ElseIf HRField = "ED_OMERS" Or Right(HRField, 4) = "_ORG" Or HRField = "ED_CPP" Or (HRField = "ED_OMERS_1" And glbCompSerial = "S/N - 2379W") Then
                Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
            
            ElseIf HRField = "ED_COUNTRY" Then
                Call convertCountry(OldValue)
                Call convertCountry(NewValue)
'            ElseIf HRField = getVITField("Payment Type") Then
'                Call Passing_Payment_Change_Vadim(PStatus, Date, oldValue, NewValue, xEmpNbr, PayIDs(x))
            End If
            
            'Town of Greater Napanee - Ticket #24375
            'Town of Lasalle
            'Ticket #20931 - This call is causing an issue.
            'Town of Aurora - Ticket #20931 - as per the mapping documentation
            'If glbCompSerial = "S/N - 2378W" And HRField = "ED_UIC" Then
            If (glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W") And HRField = "ED_UIC" Then
                Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
            End If
            
            'Ticket #25469 - City of Campbell River
            'Town of Greater Napanee - Ticket #24375
            'Town of Lasalle
            'Town of Aurora - Ticket #20931 - as per the mapping documentation
            If (glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W") And HRField = "ED_GROSSCD" Then
                Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
            End If
            
            'City of Kawartha Lakes - Ticket #21820
            If glbCompSerial = "S/N - 2363W" And HRField = "ED_GROSSCD" Then
                Call Passing_PAYINFO_Vadim(PayIDs(X), HRField, OldValue, NewValue)
            End If
            
            'City of Niagara Falls  - Get the Admin By Code to pass from Code Matrix
            'Ticket #20053 - Changed from Admin By to Benefit Group Code
            If glbCompSerial = "S/N - 2276W" And HRField = "EMP_CLASS_CODE" Then
                NewValue = CodeMatrix("BGMF", CStr(NewValue), "")
                OldValue = CodeMatrix("BGMF", CStr(OldValue), "")
            End If
            
            'Town of Aurora - Ticket #20931
            If glbCompSerial = "S/N - 2378W" And (HRField = "ED_VADIM1" Or HRField = "ED_VADIM2") Then
                NewValue = 0
                OldValue = 0
            End If
            
            If HRField = "JH_DHRS" Then
                If NewValue = "" Then NewValue = 0
            End If
            
            'Ticket #23795 - Town of Lasalle - Transfer DOH if First Day is blank and DOH has changed
            If glbCompSerial = "S/N - 2379W" And HRField = "ED_FDAY" Then
                If Not IsDate(NewValue) Then
                    NewValue = GetEmpData(xEmpnbr, "ED_DOH")
                End If
            End If
                        
            'Ticket #24996 - City of Campbell River
            If glbCompSerial = "S/N - 2458W" And HRField = "JB_DESCR" Then
                'Transfer the Organization 1 Code Desc as the Job Title
                NewValue = GetTABLDesc("ORGN", NewValue)
            End If
                
            'General Transfer -----------------------------------------------------------------
            If glbCompSerial = "S/N - 2363W" And HRField = "ED_VADIM1" Then 'City of Kawartha Lakes
                'do nothing
            
            'City of Timmins or City of Niagara Falls
            ElseIf (glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W") And HRField = "ED_VADIM2" Then
                'do nothing
            Else
                Call VadimInterface(xBatchID, PayIDs(X), HRField, OldValue, NewValue)
            End If
            '----------------------------------------------------------------------------------
            
            'City of Timmins
            If glbCompSerial = "S/N - 2375W" Then
                If HRField = "ED_PENPCT" Then
                    Call VadimInterface(xBatchID, PayIDs(X), "ED_VADIM2", OldValue, NewValue)
                End If
                If HRField = "ED_PENSION" Then
                    If NewValue = "Nu" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", OldValue, "Null")
                    Else
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", OldValue, NewValue)
                    End If
                End If
                'If HRField = "ED_OMERS" Then
                '    Call VadimInterface(xBatchID, PayIDs(x), "ED_OMERS_1", oldValue, NewValue)
                'End If
                'If HRField = "ED_NORMALR" Then
                '    If DateDiff("yyyy", oldValue, NewValue) >= 60 Then
                '        Call VadimInterface(xBatchID, PayIDs(x), "ED_OMERS_1", 2, 2)
                '    Else
                '        Call VadimInterface(xBatchID, PayIDs(x), "ED_OMERS_1", 1, 1)
                '    End If
                'End If
                If HRField = "ED_VADIM1" Then
                    xPayType = GetEmpData(xEmpnbr, "ED_REGION")
                    xDept = GetEmpData(xEmpnbr, "ED_DEPTNO")
                    If xPayType = "H" And (Val(xDept) >= 4402 And Val(xDept) <= 4412) Then
                        Call VadimInterface(xBatchID, PayIDs(X), HRField, OldValue, NewValue)
                    End If
                End If
                
                If glbCompSerial = "S/N - 2375W" Then
                    If HRField = "ED_VACPC" Then
                        If Val(NewValue) > 0 Then
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "Y")
                        Else
                            If PStatus = Banking Then
                                Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "N")
                            End If
                        End If
                    End If
                End If
                
                If HRField = "JH_EMP" Then
                    If NewValue = "NONA" Then   'Not Active Employment Status
                        '- Only transfer if Inactive and they will change manually in Vadim to Active.
                        Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_ACTIVE_FLAG:Y", "", "N")
                    'Else
                    '    Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_ACTIVE_FLAG:Y", "", "Y")
                    End If
                End If
            Else
                'Town of Aurora
                If glbCompSerial = "S/N - 2378W" Then
                    If HRField = "ED_VACPC" Then
                        If NewValue = "0.04" Then
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "Y")
                        Else
                            If PStatus = Banking Then
                                Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "N")
                            End If
                        End If
                    End If
                    
                    'Ticket #20931 - as per the mapping documentation
                    If HRField = "ED_VADIM1" Or HRField = "ED_VADIM2" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "RESET:BF_BCODE", OldValue, NewValue)
                    End If
                End If

                'And not City of Niagara Falls, Town of Lasalle, Town of Greater Napanee
                If HRField = "ED_OMERS" And glbCompSerial <> "S/N - 2276W" And glbCompSerial <> "S/N - 2379W" And glbCompSerial <> "S/N - 2447W" Then
                    Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", OldValue, NewValue)
                End If
            End If
            
            'City of Kawartha Lakes
            If glbCompSerial = "S/N - 2363W" Then
                'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
                'If HRField = "ED_ORG" Or HRField = "JH_ORG" Then
                '    If NewValue = "3" Then xCompNo = "2" Else xCompNo = "1"
                '    'If oldValue = "3" Then xOCompNo = "2" Else xOCompNo = "1"
                '    Call VadimInterface(xBatchID, PayIDs(X), "COMP_CODE", xOCompNo, xCompNo)
                'Else
                If HRField = "ED_PENSION" Then
                    Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", OldValue, NewValue)
                End If
                If HRField = "JB_ORG" Or HRField = "ED_ORG" Or HRField = "JH_ORG" Then
                    If NewValue <> "1" And NewValue <> "2" And NewValue <> "N" Then
                        'Call VadimInterface(xBatchID, PayIDs(x), "PROBATION_DATE", "", "")
                        Call VadimInterface(xBatchID, PayIDs(X), "SH_GRADE", "", "")
                        Call VadimInterface(xBatchID, PayIDs(X), "SH_GRADE_SAME", "", "")
                    End If
                End If
                If HRField = "ED_LOC" Then
                    If NewValue = "L" Or OldValue = "L" Then
                        NewValue = GetEmpData_PayrollID(PayIDs(X), "ED_LDAY")
                        Call VadimInterface(xBatchID, PayIDs(X), "TERM:LO-DATE", "", NewValue)
                    End If
                                        
                    'if PayType(Vadim) is F then update Volunteer Firefighter Flag in Vadim to Yes
                    If NewValue = "F" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "DFLT:VOL_FIREMAN_FLAG:N", "", "Y")
                    'ElseIf oldValue = "F" Then  'Ticket #20007 - Do not pass N to VFF
                    '    Call VadimInterface(xBatchID, PayIDs(X), "DFLT:VOL_FIREMAN_FLAG:N", "", "N")
                    End If
                End If
                If HRField = "ED_VACPC" And PStatus <> Demographices Then
                    If Val(NewValue) > 0 And GetEmpData(xEmpnbr, "ED_ORG") <> "3" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "Y")
                    Else
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "N")
                    End If
                ElseIf HRField = "JB_ORG" Or HRField = "ED_ORG" Or HRField = "JH_ORG" Then
                    If NewValue <> "3" And GetEmpData(xEmpnbr, "ED_VACPC") > 0 Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "Y")
                    Else
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "N")
                    End If
                End If
            End If
            
            'City of Niagara Falls
            If glbCompSerial = "S/N - 2276W" Then
            
                If HRField = "JH_EMP" Then
                    If NewValue = "VOL" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "DFLT:VOL_FIREMAN_FLAG:N", "", "Y")
                    ElseIf OldValue = "VOL" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "DFLT:VOL_FIREMAN_FLAG:N", "", "N")
                    End If
                End If
                If HRField = "ED_OMERS" Then    'RPP #
                    If NewValue = "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "")
                    Else
                        xTmpValue = GetEmpData_PayrollID(PayIDs(X), "ED_EMP")
                        'If (xTmpValue = "" Or IsNull(xTmpValue)) And UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
                        If UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
                            xTmpValue = frmEESTATS.clpCode(1).Text
                        End If
                        If xTmpValue = "FIRE" Then
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "2")
                        Else
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "1")
                        End If
                        
                    End If
                End If
                If HRField = "JH_EMP" Then    'RPP #
                    xTmpValue = GetEmpData_PayrollID(PayIDs(X), "ED_OMERS")
                    If UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
                        xTmpValue = frmEESTATS.dlpDate(2).Text
                    End If
                    If NewValue = "FIRE" And xTmpValue <> "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "2")
                    ElseIf xTmpValue <> "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "1")
                    ElseIf xTmpValue = "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "")
                    End If
                End If
                
                'Only Niagara Falls
                If glbCompSerial = "S/N - 2276W" Then
                    If HRField = "ED_ORG" Or HRField = "JH_ORG" Or HRField = "ED_PT" Or HRField = "ED_EMP" Or HRField = "ED_DEPTNO" Or HRField = "JH_DEPTNO" Or HRField = "ED_LOC" Or HRField = "ED_REGION" Or HRField = "ED_ADMINBY" Or HRField = "ED_SECTION" Then     'Union, Category or Employee Status,Location,Region,Admin By, Section
                        'Update Maximum Bank Hours in Vadim
                        'Get the OM_MAX_BANK_HRS
                        'And do not tranfer for FIRE with Dept 2850 - Ticket #14285
                        If UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
                            xUnion = frmEESTATS.clpCode(2).Text
                        Else
                            xUnion = GetEmpData(xEmpnbr, "ED_ORG")
                        End If
                        If xUnion = "FIRE" Then
                            'Get Dept No
                            If UCase(MDIMain.ActiveForm.name) = "FRMEEBASIC" Then
                                xDept = frmEEBASIC.clpDept.Text
                            Else
                                xDept = GetEmpData(xEmpnbr, "ED_DEPTNO")
                            End If
                            'Do not pass Max Bank Hours if Dept Name has "FIRE SUPPRESSION" in it.
                            xDeptName = GetDeptName(xDept, "DF_NAME")
                            If InStr(xDeptName, "FIRE SUPPRESSION") = 0 Then
                                NewValue = Get_Maximum_Bank_Hours(xEmpnbr)
                                Call VadimInterface(xBatchID, PayIDs(X), "ED_SUPCODE", "", NewValue)
                            End If
                        Else
                            NewValue = Get_Maximum_Bank_Hours(xEmpnbr)
                            If NewValue = "" And NewHireForms.count > 0 Then
                                'Do not export anything
                            Else
                                Call VadimInterface(xBatchID, PayIDs(X), "ED_SUPCODE", "", NewValue)
                            End If
                        End If
                    End If
                End If
            End If
                        
            'Dist. of Muskoka
            If glbCompSerial = "S/N - 2373W" Then
                If HRField = "ED_MSTAT" Then
                    Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_CLASS_CODE:1", OldValue, NewValue)
                End If
'                If HRField = "ED_OMERS" Then    'RPP #
'                    If Not IsDate(NewValue) Then
'                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "")
'                    Else
'                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "1")
'                    End If
'                End If
                If HRField = "ED_VACPC" Then
                    If Val(NewValue) >= 0.04 Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "Y")
                    Else
                        If PStatus = Banking Then
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "N")
                        End If
                    End If
                End If
            End If
            
            'Ticket #23795 - Town of Lasalle
            If glbCompSerial = "S/N - 2379W" Then
                'OMERS based on Pension Code
                If HRField = "ED_OMERS_1" Then
                    'OMERS
                    'If NewValue = "5" Or NewValue = "6" Then
                    If NewValue = "1" Or NewValue = "2" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS", "", "Y")
                    Else
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS", "", "N")
                    End If
                End If
                                
                'Set Pay Every Pay Period to YES when Vacation % > 0
                If HRField = "ED_VACPC" Then
                    If Val(NewValue) > 0 Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "Y")
                    Else
                        If PStatus = Banking Then
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "N")
                        End If
                    End If
                End If
            End If
            
            'Ticket #24996 - City of Campbell River - Do not transfer when New Hire as it duplicates the fields.
            If glbCompSerial = "S/N - 2458W" Then
                If HRField = "JH_PAYROLL_CATEGORY" And NewHireForms.count <= 0 Then
                    Call VadimInterface(xBatchID, PayIDs(X), "RESET:ED_EMPTYPE", OldValue, NewValue)
                End If
                If HRField = getVITField("EI Start Date") Then
                    Call VadimInterface(xBatchID, PayIDs(X), "VIT:EI Start Date", OldValue, NewValue)
                End If
                If HRField = "EMP_DEFAULT_JOB" Then
                    Call VadimInterface(xBatchID, PayIDs(X), "JH_JOB_1", OldValue, NewValue)
                End If
            End If
            
            'Ticket #25412 - Town of Greater Napanee
            If glbCompSerial = "S/N - 2447W" Then
                If HRField = "ED_OMERS" Then    'RPP #
                    If NewValue = "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "")
                    Else
                        xTmpValue = GetEmpData_PayrollID(PayIDs(X), "ED_PT")
                        If UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
                            xTmpValue = frmEESTATS.clpPT.Text
                        End If
                        'Get Dept No
                        If UCase(MDIMain.ActiveForm.name) = "FRMEEBASIC" Then
                            xDept = frmEEBASIC.clpDept.Text
                        Else
                            xDept = GetEmpData(xEmpnbr, "ED_DEPTNO")
                        End If
                        If xTmpValue = "FT" And xDept = "2100" Then
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "2")
                        Else
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "1")
                        End If
                        
                    End If
                End If
                If HRField = "JH_PT" Or HRField = "ED_PT" Or HRField = "JH_PAYROLL_CATEGORY" Then      'RPP #
                    xTmpValue = GetEmpData_PayrollID(PayIDs(X), "ED_OMERS")
                    If UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
                        xTmpValue = frmEESTATS.dlpDate(2).Text
                    End If
                    'Get Dept No
                    If UCase(MDIMain.ActiveForm.name) = "FRMEEBASIC" Then
                        xDept = frmEEBASIC.clpDept.Text
                    Else
                        xDept = GetEmpData(xEmpnbr, "ED_DEPTNO")
                    End If
                    If NewValue = "FT" And xDept = "2100" And xTmpValue <> "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "2")
                    ElseIf xTmpValue <> "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "1")
                    ElseIf xTmpValue = "" Then
                        Call VadimInterface(xBatchID, PayIDs(X), "ED_OMERS_1", "", "")
                    End If
                    
                    'Ticket #30112 - Do not transfer PAY_VAC_PER_FLAG on New Hire. There are # of fields not to transfer on New Hire and this one of them. Otherwise, the
                    'EMPLOYEE record does not get added in Vadim.
                    'Ticket #29727 - If Category = PT set Pay Every Period to YES.
                    If (HRField = "JH_PT" Or HRField = "ED_PT" Or HRField = "JH_PAYROLL_CATEGORY") And NewHireForms.count <= 0 Then
                        If NewValue = "PT" Then
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "Y")
                        Else
                            Call VadimInterface(xBatchID, PayIDs(X), "ED_VACPC_1", "N", "N")
                        End If
                    End If
                End If
            End If
            
            If HRField = "SH_GRADE" Then
                'City of Timmins
                If glbCompSerial = "S/N - 2375W" Then
                    'If Val(oldValue) < Val(NewValue) Then
                        Call VadimInterface(xBatchID, PayIDs(X), "SH_GRADE_SAME", OldValue, NewValue)
                    'End If
                Else
                    'City of Kawartha Lakes
                    If glbCompSerial = "S/N - 2363W" Then
                        'Get Pay Type, Only populate After Probation Lvl and Prob Date if Pay Type=H
                        xPayType = GetEmpData(xEmpnbr, "ED_LOC")
                        If xPayType = "H" Then
                            'Get Next Level
                            xNxtGrdFlg = Next_Step_Available(PayIDs(X), NewValue)
                            If xNxtGrdFlg = True Then
                                Call VadimInterface(xBatchID, PayIDs(X), "SH_GRADE_SAME", "", Val(NewValue) + 1)
                            Else
                                Call VadimInterface(xBatchID, PayIDs(X), "SH_GRADE_SAME", "", NewValue)
                                'Call VadimInterface(xBatchID, PayIDs(x), "PROBATION_DATE", "", "01/01/2099")
                            End If
                        Else
                            Call VadimInterface(xBatchID, PayIDs(X), "SH_GRADE_SAME", "", "")
                            Call VadimInterface(xBatchID, PayIDs(X), "DFLT:PROBATION_DATE", "", "")
                        End If
                    Else
                        'Ticket #20931 - Town of Aurora uses this update as well.
                        Call VadimInterface(xBatchID, PayIDs(X), "SH_GRADE_SAME", OldValue, NewValue)
                    End If
                End If
            End If
        End If
    Next Y
    'end adding information to the interface detail table
    
    'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
    'transfer. Vadim has logic behind the Process Code P.
    If glbCompSerial = "S/N - 2373W" And PStatus = "4" And (CDate(UptDate) > CDate(Date)) Then
        Call CloseBatchVadim(xBatchID, "P") ' change the process code from I to P in the interface header table
    Else
        Call CloseBatchVadim(xBatchID) ' change the process code from I to N in the interface header table
    End If
    
    'City of Niagara Falls - Ticket #15542
    'If Salary changes and City of Niagara Falls then update info:HR HR_VADIM_SY_INTERFACE table
    'to track the salary changes taking places to enable future undoing of the salary change.
    If glbCompSerial = "S/N - 2276W" And PStatus = "4" And (CDate(UptDate) > CDate(Date)) And UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then      'Salary changes
        If frmESALARY.txtVadPayRate.Text <> "" Then
            Call Update_HR_Vadim_Sy_Interface(xBatchID, xPayID, UptDate, frmESALARY.txtVadPayRate.Text, GetJHData(xPayID, "JH_JOB", ""), frmESALARY.txtVadAddModDel.Text, Format(Now, "mm/dd/yyyy h:m:s"), "Pay change info.")
        End If
        If frmESALARY.txtVadSalRate.Text <> "" Then
            Call Update_HR_Vadim_Sy_Interface(xBatchID, xPayID, UptDate, frmESALARY.txtVadSalRate.Text, GetJHData(xPayID, "JH_JOB", ""), frmESALARY.txtVadAddModDel.Text, Format(Now, "mm/dd/yyyy h:m:s"), "Salary change info.")
        End If
    End If
        
Next X

Exit Sub

Passing_Bank_Changes_Vadim_Err:
MsgBox Err.Description
Resume Next

End Sub

Sub Passing_Attendance_Vadim(HRAtts, UptType, xEmpnbr, Optional xPayID)
'NOT IN USE IN VB CODE. HAS TRIGGER IN DATABASE
'Dim x, xBNo
'Dim PayIDs
'Dim VDClause
'Dim oldReason, NewReason
'Dim oldPayCode, newPayCode
'Dim oldDate, NewDate
'Dim oldHours, NewHours
'Dim oldJob, newJob
'Dim oldGLNo, newGLNo
'Dim oldSalary, newSalary
'Dim oldSalCode, newSalCode
'Dim oldRate, newRate
'Dim oldWhrs, newWhrs
'Dim oldComment, newComment
'Dim oldAmount, newAmount
'Dim rsEmp As New ADODB.Recordset
'Dim PayCodeInfo As PayCodeInfoType
'Dim AttBatchID
'On Error GoTo Passing_Attendance_Vadim_Err
'If Not isTransfer(Attendance) Then Exit Sub
'
'
'PayIDs = Split(getPayrollIDs(xEmpnbr, xPayID, True), "|")
'For x = 0 To UBound(PayIDs)
'    xPayID = PayIDs(x)
'    oldReason = HRAtts("AD_REASON").OldValue
'    NewReason = HRAtts("AD_REASON").NewValue
'
'    Call getPayCodeInfo(PayCodeInfo, "ATTENDANCE", oldReason)
'    oldPayCode = PayCodeInfo.PayCode
'    Call getPayCodeInfo(PayCodeInfo, "ATTENDANCE", NewReason)
'    newPayCode = PayCodeInfo.PayCode
'    oldDate = HRAtts("AD_DOA").OldValue
'    NewDate = HRAtts("AD_DOA").NewValue
'    oldHours = HRAtts("AD_HRS").OldValue
'    NewHours = HRAtts("AD_HRS").NewValue
'
'    If UptType = "D" Then
'        VDClause = "PAY_CODE='" & newPayCode & "' AND TRANS_DATE='" & Date_SQL(NewDate) & " AND TRANS_HOURS=" & NewHours
'        AttBatchID = AddBatchVadim("D", Date)
'        Call VadimInterface(AttBatchID, xPayID, "TRANS:PAYCODE", newPayCode, "")
'        Call CloseBatchVadim(AttBatchID)
'    Else
'        oldJob = HRAtts("AD_JOB").OldValue
'        newJob = HRAtts("AD_JOB").NewValue
'        oldGLNo = getGLNo(xEmpnbr, oldJob)
'        newGLNo = getGLNo(xEmpnbr, newJob)
'
'        oldSalary = HRAtts("AD_SALARY").OldValue
'        newSalary = HRAtts("AD_SALARY").NewValue
'        oldSalCode = HRAtts("AD_SALCD").OldValue
'        newSalCode = HRAtts("AD_SALCD").NewValue
'        oldWhrs = HRAtts("AD_WHRS").OldValue
'        newWhrs = HRAtts("AD_WHRS").NewValue
'        oldComment = HRAtts("AD_COMM").OldValue
'        newComment = HRAtts("AD_COMM").NewValue
'        oldRate = 0
'        If oldSalCode = "A" Then
'            If oldWhrs <> 0 Then
'                oldRate = CStr(Round(oldSalary / oldWhrs / 52, 2))
'            End If
'        ElseIf oldSalCode = "H" Then
'            oldRate = oldSalary
'        End If
'        oldAmount = CStr(Round(oldRate * oldHours, 2))
'        newRate = 0
'        If newSalCode = "A" Then
'            If newWhrs <> 0 Then
'                newRate = CStr(Round(newSalary / newWhrs / 52, 2))
'            End If
'        ElseIf newSalCode = "H" Then
'            newRate = newSalary
'        End If
'        newAmount = CStr(Round(newRate * NewHours, 2))
'
'        If UptType = "A" Then
'            AttBatchID = AddBatchVadim("A", NewDate)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:PAYROLL_ID", Null, xPayID)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:GLNO", Null, newGLNo)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:JOBCODE", Null, newJob)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:PAYCODE", Null, newPayCode)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:DATE", Null, NewDate)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:CREATE_TYPE", Null, "M")
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:HOURS", Null, NewHours)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:RATE", Null, newRate)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:AMOUNT", Null, newAmount)
'            Call VadimInterface(AttBatchID, xPayID, "TRANS:DESC", Null, newComment)
'            Call CloseBatchVadim(AttBatchID)
'        Else
'            VDClause = "PAY_CODE='" & newPayCode & "' AND TRANS_DATE='" & Date_SQL(NewDate) & " AND TRANS_HOURS=" & NewHours
'            AttBatchID = AddBatchVadim("M", NewDate)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:GLNO", oldGLNo, newGLNo)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:JOBCODE", oldJob, newJob)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:PAYCODE", oldPayCode, newPayCode)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:DATE", oldDate, NewDate)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:HOURS", oldHours, NewHours)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:RATE", oldRate, newRate)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:AMOUNT", oldAmount, newAmount)
''            Call VadimInterface(AttBatchID, xPayID, "TRANS:DESC", oldComment, newComment)
'
'            If oldGLNo <> newGLNo Then Call VadimInterface(AttBatchID, xPayID, "TRANS:GLNO", oldGLNo, newGLNo)
'            If oldJob <> newJob Then Call VadimInterface(AttBatchID, xPayID, "TRANS:JOBCODE", oldJob, newJob)
'            If oldPayCode <> newPayCode Then Call VadimInterface(AttBatchID, xPayID, "TRANS:PAYCODE", oldPayCode, newPayCode)
'            If oldDate <> NewDate Then Call VadimInterface(AttBatchID, xPayID, "TRANS:DATE", oldDate, NewDate)
'            If oldHours <> NewHours Then Call VadimInterface(AttBatchID, xPayID, "TRANS:HOURS", oldHours, NewHours)
'            If oldRate <> newRate Then Call VadimInterface(AttBatchID, xPayID, "TRANS:RATE", oldRate, newRate)
'            If oldAmount <> newAmount Then Call VadimInterface(AttBatchID, xPayID, "TRANS:AMOUNT", oldAmount, newAmount)
'            If oldComment <> newComment Then Call VadimInterface(AttBatchID, xPayID, "TRANS:DESC", oldComment, newComment)
'
'            Call CloseBatchVadim(AttBatchID)
'         End If
'    End If
'Next x
'Exit Sub
'Passing_Attendance_Vadim_Err:
'MsgBox Err.Description
'Resume Next
End Sub
Sub Passing_Payment_Change_Vadim(PStatus As PassStatus, UptDate, oldPayType, newPayType, xEmpnbr, xPayID)
''If Not IsMissing(newPayType) Then
''    If (oldPayType = "P" Or oldPayType = "H" Or oldPayType = "F" Or oldPayType = "C") Then oldPayType = "H"
''    If (newPayType = "P" Or newPayType = "H" Or newPayType = "F" Or newPayType = "C") Then newPayType = "H"
''    If (oldPayType = "S" And newPayType = "H") Then
''
''    End If
''
''    If (oldPayType = "H" And newPayType = "S") Then
''    End If
''    Exit Sub
''End If
End Sub

Sub Passing_Salary_Vadim(HRSalary As Collection, PStatus As PassStatus, UptDate, xPHrs, xWHRS, xEmpnbr, xPayID, Optional newPayType, Optional xNiagaraWHRS)
Dim X
Dim VDClause
Dim UptType
Dim PayCodeInfo As PayCodeInfoType
Dim SalaryBatchID
Dim oldSalary
Dim newSalary
Dim oldSalCD
Dim newSalCD
Dim oldRate, newRate
Dim Salary_PayCode, Rate_PayCode
Dim Salary_PayFreq, Rate_PayFreq
Dim xCompNo, xUnion As String
Dim xPayPeriod
Dim locWHRS

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(Salary) Then Exit Sub

'Ticket #24565 - District Municipality of Muskoka
'NEW 10/22/2014: This is getting commented out as they do not want the Occupation Rate to be updated for 181W.
'Since we are transferring Occupation Code/Job Group is Employee # for Employee in Union '181W', the
'Occupation record needs to be updated with the new Total Rate as Hourly Rate of the Occupation Code when employee's
'Salary changes
'If glbCompSerial = "S/N - 2373W" And GetEmpData_PayrollID(xEMPNBR, "ED_ORG") = "181W" Then
'    Call Passing_OccupationData_Vadim(xEMPNBR, HRSalary("SH_SALARY").OldValue, HRSalary("SH_SALARY").NewValue, UptDate)
'End If

If IsMissing(newPayType) Then
    newPayType = getPayType(xEmpnbr)
End If

oldSalCD = HRSalary("SH_SALCD").OldValue
newSalCD = HRSalary("SH_SALCD").NewValue

oldRate = 0
oldSalary = 0
If oldSalCD = "H" Then
    'City of Niagara Falls - Round the Hourly Rate to 2 Decimal places - Hourly to Hourly Rate conversion
    If glbCompSerial = "S/N - 2276W" Then
        oldRate = Round(HRSalary("SH_SALARY").OldValue, 4) 'round to 4 decimal places instead of 2
        'oldSalary = Round2DEC(oldRate * Val(frmESALARY.txtWHRS.Text))
        'Ticket #19064
        If IsMissing(xNiagaraWHRS) Then
            If UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
                oldSalary = Round(oldRate * Val(frmESALARY.txtWHRS.Text), 4) 'round to 4 decimal places
            Else
                locWHRS = GetSHData(xEmpnbr, "SH_WHRS", 0)
                oldSalary = Round(oldRate * Val(locWHRS), 4) 'round to 4 decimal places
            End If
        Else
            oldSalary = Round(oldRate * xNiagaraWHRS, 4) 'round to 4 decimal places
        End If
    Else
        If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
            oldRate = HRSalary("SH_TOTAL").OldValue
            oldSalary = Round2DEC(oldRate * xPHrs)
        Else
            oldRate = HRSalary("SH_SALARY").OldValue
            oldSalary = Round2DEC(oldRate * xPHrs)
        End If
    End If
    
ElseIf oldSalCD = "A" Then
    'City of Niagara Falls - Special formula to calculate Hourly Rate - xWHRS actually contains Hours per Day
    If glbCompSerial = "S/N - 2276W" Then
        'Round the Hourly Rate to 4 Decimal places - Salary to Hourly Rate conversion
        'If xWHRS <> 0 Then oldRate = Round((HRSalary("SH_SALARY").oldValue / xPHrs) / (xWHRS * 5), 4)
        'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
        'So xPHrs contains Pay Periods per Year (SH_PAYP) and xWHRS contains Hours Per Pay (JB_DHRS)
        If xWHRS <> 0 Then oldRate = Round((HRSalary("SH_SALARY").OldValue / xPHrs / xWHRS), 4)
        
        'oldSalary = Round2DEC(oldRate * Val(frmESALARY.txtWHRS.Text))
        'oldSalary = Round(oldRate * Val(frmESALARY.txtWHRS.Text), 4) 'round to 4 decimal places
        'Ticket #16276 - Same formula as on the Salary screen
        'Ticket #18257 - Provided the formula to calculate the Salary per Pay - using only # of Pay Periods
        'oldSalary = Round(((HRSalary("SH_SALARY").oldValue / 52) / Val(frmESALARY.txtWHRS.Text)) * xWHRS, 4)
        oldSalary = Round((HRSalary("SH_SALARY").OldValue / xPHrs), 4)
    Else
        If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
            If xWHRS <> 0 Then oldRate = Round2DEC((HRSalary("SH_TOTAL").OldValue / 52) / xWHRS)
            'oldSalary = Round2DEC(oldRate * xPHrs)
            If xWHRS <> 0 Then oldSalary = Round2DEC(((HRSalary("SH_TOTAL").OldValue / 52) / xWHRS) * xPHrs)
        Else
            If xWHRS <> 0 Then oldRate = Round2DEC((HRSalary("SH_SALARY").OldValue / 52) / xWHRS)
            'oldSalary = Round2DEC(oldRate * xPHrs)
            If xWHRS <> 0 Then oldSalary = Round2DEC(((HRSalary("SH_SALARY").OldValue / 52) / xWHRS) * xPHrs)
        End If
    End If
ElseIf oldSalCD = "M" Then
    If xWHRS <> 0 Then oldRate = Round2DEC(((HRSalary("SH_SALARY").OldValue * 12) / 52) / xWHRS)
    oldSalary = Round2DEC(oldRate * xPHrs)
End If

newRate = 0
newSalary = 0
If newSalCD = "H" Then
    'City of Niagara Falls - Round the Hourly Rate to 2 Decimal places - Hourly to Hourly Rate conversion
    If glbCompSerial = "S/N - 2276W" Then
        newRate = Round(HRSalary("SH_SALARY").NewValue, 4)  'round to 4 decimal places instead of 2
        'newSalary = Round2DEC(newRate * Val(frmESALARY.txtWHRS.Text))
        'Ticket #19064
        If IsMissing(xNiagaraWHRS) Then
            If UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
                newSalary = Round(newRate * Val(frmESALARY.txtWHRS.Text), 4) 'round to 4 decimal places
            Else
                locWHRS = GetSHData(xEmpnbr, "SH_WHRS", 0)
                newSalary = Round(newRate * Val(locWHRS), 4) 'round to 4 decimal places
            End If
        Else
            newSalary = Round(newRate * xNiagaraWHRS, 4) 'round to 4 decimal places
        End If
    Else
        If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
            newRate = HRSalary("SH_TOTAL").NewValue
            newSalary = Round2DEC(newRate * xPHrs)
        Else
            newRate = HRSalary("SH_SALARY").NewValue
            newSalary = Round2DEC(newRate * xPHrs)
        End If
    End If
    
ElseIf newSalCD = "A" Then
    'City of Niagara Falls - Special formula to calculate Hourly Rate - xWHRS actually contains Hours per Day
    If glbCompSerial = "S/N - 2276W" Then
        'Round the Hourly Rate to 4 Decimal places - Salary to Hourly Rate conversion
        'If xWHRS <> 0 Then newRate = Round((HRSalary("SH_SALARY").NewValue / xPHrs) / (xWHRS * 5), 4)
        'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
        'So xPHrs contains Pay Periods per Year (SH_PAYP) and xWHRS contains Hours Per Pay (JB_DHRS)
        If xWHRS <> 0 Then newRate = Round((HRSalary("SH_SALARY").NewValue / xPHrs / xWHRS), 4)
        
        'newSalary = Round2DEC(newRate * Val(frmESALARY.txtWHRS.Text))
        'newSalary = Round(newRate * Val(frmESALARY.txtWHRS.Text), 4) 'round to 4 decimal places
        'Ticket #16276 - Same formula as on the Salary screen
        'Ticket #18257 - Provided the formula to calculate the Salary per Pay - using only # of Pay Periods
        'newSalary = Round(((HRSalary("SH_SALARY").NewValue / 52) / Val(frmESALARY.txtWHRS.Text)) * xWHRS, 4)
        newSalary = Round((HRSalary("SH_SALARY").NewValue / xPHrs), 4)
    Else
        If glbCompSerial = "S/N - 2373W" Then   'DMuskoka  - Pass Total which includes Premium
            If xWHRS <> 0 Then newRate = Round2DEC((HRSalary("SH_TOTAL").NewValue / 52) / xWHRS)
            'newSalary = Round2DEC(newRate * xPHrs)
            If xWHRS <> 0 Then newSalary = Round2DEC(((HRSalary("SH_TOTAL").NewValue / 52) / xWHRS) * xPHrs)
        Else
            If xWHRS <> 0 Then newRate = Round2DEC((HRSalary("SH_SALARY").NewValue / 52) / xWHRS)
            'newSalary = Round2DEC(newRate * xPHrs)
            If xWHRS <> 0 Then newSalary = Round2DEC(((HRSalary("SH_SALARY").NewValue / 52) / xWHRS) * xPHrs)
        End If
    End If
ElseIf newSalCD = "M" Then
    If xWHRS <> 0 Then newRate = Round2DEC(((HRSalary("SH_SALARY").NewValue * 12) / 52) / xWHRS)
    newSalary = Round2DEC(newRate * xPHrs)
End If

UptType = "M"
If oldSalCD = "" And newSalCD <> "" Then UptType = "A"
If oldSalCD <> "" And newSalCD = "" Then UptType = "D"
    
Call getPayCodeInfo(PayCodeInfo, "RATE")
If Len(PayCodeInfo.PayCode) = 0 Then Exit Sub
If Len(PayCodeInfo.PayFreq) = 0 Then Exit Sub
Rate_PayCode = PayCodeInfo.PayCode
Rate_PayFreq = PayCodeInfo.PayFreq


Call getPayCodeInfo(PayCodeInfo, "SALARY")
If Len(PayCodeInfo.PayCode) = 0 Then Exit Sub
If Len(PayCodeInfo.PayFreq) = 0 Then Exit Sub
Salary_PayCode = PayCodeInfo.PayCode
Salary_PayFreq = PayCodeInfo.PayFreq

Select Case UptType
Case "A", "M"
    'For District Municipality of Muskoka - for Councilors do not transfer Rate. Only transfer Salary
    'Ticket #19113
    If glbCompSerial = "S/N - 2373W" And GetEmpData(xEmpnbr, "ED_EMPTYPE") = "9" Then
        GoTo only_Salary_Transfer
    'ElseIf glbCompSerial = "S/N - 2379W" And newPayType = "S" Then
    '    'Town of Lasalle - Only Salary Transfer for PayType = 'S'
    '    GoTo only_Salary_Transfer
    End If
    
    If isExistTransCode(xPayID, Rate_PayCode) = 1 Then
        SalaryBatchID = AddBatchVadim("M", UptDate)
        VDClause = "PAY_CODE='" & Rate_PayCode & "'"
        
        'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
        'City of Kawartha Lakes - Company Code
        'If glbCompSerial = "S/N - 2363W" Then
        '    xUnion = GetEmpData(xEMPNBR, "ED_ORG")
        '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
        '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
        '    Call VadimInterface(SalaryBatchID, xPayid, "TRANSCODE:COMP_CODE", xCompNo, xCompNo, VDClause)
        'End If
        
        Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYCODE", Rate_PayCode, Rate_PayCode, VDClause)
        Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:FREQCODE", Rate_PayFreq, Rate_PayFreq, VDClause)
        
        'Ticket #24565 - District Municipality of Muskoka
        'NEW 10/22/2014: They want the Trans Value (TOTAL) to transfer as they are now maintaining the rates in info:HR
        'They don't want Trans Value to transfer for anyone, only transfer the Pay Code itself and other values
        'If glbCompSerial = "S/N - 2373W" Then
        '    'Do not transfer Trans Value
        '    'They want to test with 0 value transfer for Trans Value
        '    Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, 0, VDClause)
        'Else
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", oldRate, newRate, VDClause)
        'End If
        
        'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
        'transfer. Vadim has logic behind the Process Code P.
        If glbCompSerial = "S/N - 2373W" And (CDate(UptDate) > CDate(Date)) Then
            Call CloseBatchVadim(SalaryBatchID, "P") ' change the process code from I to P in the interface header table
        Else
            Call CloseBatchVadim(SalaryBatchID)
        End If
        
        'City of Niagara Falls - Ticket #15542
        'Update HR_VADIM_SY_INTERFACE table with this entry - for keeping track of Salary or
        'related information changes for future maintenance purpose.
        If glbCompSerial = "S/N - 2276W" And (CDate(UptDate) > CDate(Date)) Then
            If IsMissing(xNiagaraWHRS) Then
                If UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
                    frmESALARY.txtVadPayRate.Text = newRate
                End If
            End If
            Call Update_HR_Vadim_Sy_Interface(SalaryBatchID, xPayID, UptDate, newRate, GetJHData(xPayID, "JH_JOB", ""), "M", Format(Now, "mm/dd/yyyy h:m:s"), "Pay Rate change")
        End If
    Else
        SalaryBatchID = AddBatchVadim("A", UptDate)
        
        'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
        'City of Kawartha Lakes - Company Code
        'If glbCompSerial = "S/N - 2363W" Then
        '    xUnion = GetEmpData(xEMPNBR, "ED_ORG")
        '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
        '    Call VadimInterface(SalaryBatchID, xPayid, "TRANSCODE:COMP_CODE", xCompNo, xCompNo)
        'End If
        
        Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYROLL_ID", "", xPayID)
        Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYCODE", "", Rate_PayCode)
        Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:FREQCODE", "", Rate_PayFreq)
        
        'Ticket #24565 - District Municipality of Muskoka
        'NEW 10/22/2014: They want the Trans Value (TOTAL) to transfer as they are now maintaining the rates in info:HR
        'They don't want Trans Value to transfer for anyone, only transfer the Pay Code itself and other values
        'If glbCompSerial = "S/N - 2373W" Then
        '    'Do not transfer Trans Value
        '    'They want to test with 0 value transfer for Trans Value
        '    Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, 0)
        'Else
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, newRate)
        'End If
        
        Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:EMPLR_MULTIPLIER", 0, 0)
                
        'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
        'transfer. Vadim has logic behind the Process Code P.
        If glbCompSerial = "S/N - 2373W" And (CDate(UptDate) > CDate(Date)) Then
            Call CloseBatchVadim(SalaryBatchID, "P") ' change the process code from I to P in the interface header table
        Else
            Call CloseBatchVadim(SalaryBatchID)
        End If
    
        'City of Niagara Falls - Ticket #15542
        'Update HR_VADIM_SY_INTERFACE table with this entry - for keeping track of Salary or
        'related information changes for future maintenance purpose.
        If glbCompSerial = "S/N - 2276W" And (CDate(UptDate) > CDate(Date)) Then
            If IsMissing(xNiagaraWHRS) Then
                If UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
                    frmESALARY.txtVadPayRate.Text = newRate
                End If
            End If
            Call Update_HR_Vadim_Sy_Interface(SalaryBatchID, xPayID, UptDate, newRate, GetJHData(xPayID, "JH_JOB", ""), "A", Format(Now, "mm/dd/yyyy h:m:s"), "New Pay Rate")
        End If
    End If
        
only_Salary_Transfer:

    If isExistTransCode(xPayID, Salary_PayCode) = 1 Then
        If newPayType = "S" Then
            SalaryBatchID = AddBatchVadim("M", UptDate)
            VDClause = "PAY_CODE='" & Salary_PayCode & "'"
            
            'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
            'City of Kawartha Lakes - Company Code
            'If glbCompSerial = "S/N - 2363W" Then
            '    xUnion = GetEmpData(xEMPNBR, "ED_ORG")
            '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
            '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
            '    Call VadimInterface(SalaryBatchID, xPayid, "TRANSCODE:COMP_CODE", xCompNo, xCompNo, VDClause)
            'End If
            
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYCODE", Salary_PayCode, Salary_PayCode, VDClause)
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:FREQCODE", Salary_PayFreq, Salary_PayFreq, VDClause)
            
            'Ticket #24565 - District Municipality of Muskoka
            'NEW 10/22/2014: They want the Trans Value (TOTAL) to transfer as they are now maintaining the rates in info:HR
            'They don't want Trans Value to transfer for anyone, only transfer the Pay Code itself and other values
            'If glbCompSerial = "S/N - 2373W" Then
            '    'Do not transfer Trans Value
            '    'They want to test with 0 value transfer for Trans Value
            '    Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, 0, VDClause)
            'Else
                Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", oldSalary, newSalary, VDClause)
            'End If
            
            'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
            'transfer. Vadim has logic behind the Process Code P.
            If glbCompSerial = "S/N - 2373W" And (CDate(UptDate) > CDate(Date)) Then
                Call CloseBatchVadim(SalaryBatchID, "P") ' change the process code from I to P in the interface header table
            Else
                Call CloseBatchVadim(SalaryBatchID)
            End If
            
            'City of Niagara Falls - Ticket #15542
            'Update HR_VADIM_SY_INTERFACE table with this entry - for keeping track of Salary or
            'related information changes for future maintenance purpose.
            If glbCompSerial = "S/N - 2276W" And (CDate(UptDate) > CDate(Date)) Then
                If IsMissing(xNiagaraWHRS) Then
                    If UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
                        frmESALARY.txtVadSalRate.Text = newSalary
                    End If
                End If
                Call Update_HR_Vadim_Sy_Interface(SalaryBatchID, xPayID, UptDate, newSalary, GetJHData(xPayID, "JH_JOB", ""), "M", Format(Now, "mm/dd/yyyy h:m:s"), "Salary change")
            End If
        
        ElseIf newPayType = "H" Then
            VDClause = "PAY_CODE='" & Salary_PayCode & "'"
            
            'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
            'City of Kawartha Lakes - Company Code
            'If glbCompSerial = "S/N - 2363W" Then
            '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
            'End If
            
            SalaryBatchID = AddBatchVadim("D", UptDate)
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYCODE", Salary_PayCode, "", VDClause)
            
            'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
            'transfer. Vadim has logic behind the Process Code P.
            If glbCompSerial = "S/N - 2373W" And (CDate(UptDate) > CDate(Date)) Then
                Call CloseBatchVadim(SalaryBatchID, "P") ' change the process code from I to P in the interface header table
            Else
                Call CloseBatchVadim(SalaryBatchID)
            End If
        End If
    Else
        If newPayType = "S" Then
            SalaryBatchID = AddBatchVadim("A", UptDate)
            
            'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
            'City of Kawartha Lakes - Company Code
            'If glbCompSerial = "S/N - 2363W" Then
            '    xUnion = GetEmpData(xEMPNBR, "ED_ORG")
            '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
            '    Call VadimInterface(SalaryBatchID, xPayid, "TRANSCODE:COMP_CODE", xCompNo, xCompNo)
            'End If
            
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYROLL_ID", "", xPayID)
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYCODE", "", Salary_PayCode)
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:FREQCODE", "", Salary_PayFreq)
            
            'Ticket #24565 - District Municipality of Muskoka
            'NEW 10/22/2014: They want the Trans Value (TOTAL) to transfer as they are now maintaining the rates in info:HR
            'They don't want Trans Value to transfer for anyone, only transfer the Pay Code itself and other values
            'If glbCompSerial = "S/N - 2373W" Then
            '    'Do not transfer Trans Value
            '    'They want to test with 0 value transfer for Trans Value
            '    Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, 0)
            'Else
                Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, newSalary)
            'End If
            Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:EMPLR_MULTIPLIER", 0, 0)
                        
            'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
            'transfer. Vadim has logic behind the Process Code P.
            If glbCompSerial = "S/N - 2373W" And (CDate(UptDate) > CDate(Date)) Then
                Call CloseBatchVadim(SalaryBatchID, "P") ' change the process code from I to P in the interface header table
            Else
                Call CloseBatchVadim(SalaryBatchID)
            End If
            
            'City of Niagara Falls - Ticket #15542
            'Add HR_VADIM_SY_INTERFACE table with this entry - for keeping track of Salary or
            'related information changes for future maintenance purpose.
            If glbCompSerial = "S/N - 2276W" And (CDate(UptDate) > CDate(Date)) Then
                If IsMissing(xNiagaraWHRS) Then
                    If UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
                        frmESALARY.txtVadSalRate.Text = newSalary
                    End If
                End If
                Call Update_HR_Vadim_Sy_Interface(SalaryBatchID, xPayID, UptDate, newSalary, GetJHData(xPayID, "JH_JOB", ""), "A", Format(Now, "mm/dd/yyyy h:m:s"), "New Salary")
            End If
            
        End If
    End If
Case "D"
    VDClause = "PAY_CODE='" & Rate_PayCode & "'"
    SalaryBatchID = AddBatchVadim("D", UptDate)
    
    'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
    'City of Kawartha Lakes - Company Code
    'If glbCompSerial = "S/N - 2363W" Then
    '    xUnion = GetEmpData(xEMPNBR, "ED_ORG")
    '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
    '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
    'End If
    
    Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYCODE", Rate_PayCode, "", VDClause)
    
    'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
    'transfer. Vadim has logic behind the Process Code P.
    If glbCompSerial = "S/N - 2373W" And (CDate(UptDate) > CDate(Date)) Then
        Call CloseBatchVadim(SalaryBatchID, "P") ' change the process code from I to P in the interface header table
    Else
        Call CloseBatchVadim(SalaryBatchID)
    End If
    
    If newPayType = "S" Then
        VDClause = "PAY_CODE='" & Salary_PayCode & "'"
        
        'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
        'City of Kawartha Lakes - Company Code
        'If glbCompSerial = "S/N - 2363W" Then
        '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
        'End If
        SalaryBatchID = AddBatchVadim("D", UptDate)
        Call VadimInterface(SalaryBatchID, xPayID, "TRANSCODE:PAYCODE", Salary_PayCode, "", VDClause)
                
        'Ticket #24565 - District Municipality of Muskoka - Set Process Code to P - Pending for Future Dated Salary changes
        'transfer. Vadim has logic behind the Process Code P.
        If glbCompSerial = "S/N - 2373W" And (CDate(UptDate) > CDate(Date)) Then
            Call CloseBatchVadim(SalaryBatchID, "P") ' change the process code from I to P in the interface header table
        Else
            Call CloseBatchVadim(SalaryBatchID)
        End If
    End If
End Select
End Sub

Sub Passing_Position_Master_Vadim(xOccCode, UptType, oldDesc, newDesc)
Dim X
Dim VDClause
Dim UptDate
Dim PayCodeInfo As PayCodeInfoType
Dim JobMasterBatchID
Dim OldValue
Dim NewValue
Dim xCompNo  As String
Dim xGroup As String
Dim xBaseRate
Dim xGridType As String
Dim xFTEHrsYear
Dim xAnnSal

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(PositionMaster) Then Exit Sub

UptDate = Date
If UptType = "D" Then GoTo Next_Step

If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes
    'xUnion = GetJobData(xOccCode, "JB_ORG")
    'If xUnion <> "1" And xUnion <> "2" And xUnion <> "N" Then Exit Sub
    
    'Changed the logic from above - Do not transfer Occupation Code/Position Code
    'if the Position Group <> VAD
    xGroup = GetJobData(NewValue, "JB_GRPCD")
    If xGroup = "" Then
        xGroup = frmMPOSITIONS.clpCode(2).Text
    End If
    If xGroup <> "VAD" Then Exit Sub
End If

UptDate = Date
OldValue = oldDesc
NewValue = newDesc
If IsNull(OldValue) Then OldValue = ""
If IsNull(NewValue) Then NewValue = ""

'City of Timmins - Transfer Base Rate
'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
'If glbCompSerial = "S/N - 2375W" Then
    xGridType = ""
    xBaseRate = GetJobData(xOccCode, "JB_S1")
    xGridType = GetJobData(xOccCode, "JB_SALCD")
    xFTEHrsYear = GetJobData(xOccCode, "JB_FTEHRS")
    If xBaseRate = "" Then
        xBaseRate = frmMPOSITIONS.medPayScale(1).Text
        xGridType = frmMPOSITIONS.lblSalCode.Caption
        xFTEHrsYear = frmMPOSITIONS.medFTEHrs.Text
    End If
    
    'Ticket #23795 - Town of Lasalle - New Salary Amount field to be updates if GridType is A
    xAnnSal = 0
    If xGridType = "A" Then xAnnSal = xBaseRate
    
    'Get the Hourly Rate for the Base Rate as per their special calculation
    'City of Niagara Falls
    If glbCompSerial = "S/N - 2276W" Then
        If xBaseRate <> "" Then
            xBaseRate = Get_Hourly_Rate_NiagaraFalls(xOccCode, xBaseRate, xGridType)
        Else
            xBaseRate = Get_Hourly_Rate_NiagaraFalls(xOccCode, , xGridType)
        End If
        
        'Ticket #23795 - Because CNF always wants Hourly Rate transferred. The above code computes Hourly Rate
        xGridType = "H"
    Else
        'Ticket #21124 - If Annual Grid then convert to Base Rate to Hourly Rate
        If xGridType = "A" Then
            If Val(xFTEHrsYear) <> 0 Then
                xBaseRate = Val(xBaseRate) / Val(xFTEHrsYear)
            Else
                xBaseRate = 0
            End If
        End If
    End If
'End If

If xBaseRate = "" Or IsNull(xBaseRate) Then xBaseRate = 0
If xAnnSal = "" Or IsNull(xAnnSal) Then xAnnSal = 0

'City of Timmins - Transfer Base Rate
If glbCompSerial = "S/N - 2375W" Then
    'If oldDesc = NewValue Then Exit Sub
Else
    If oldDesc = NewValue Then Exit Sub
End If

'If glbLambton Then
'    xOccCode = Left(xGrid, 1) & xJob & Mid(xGrid, 2)
'Else
'    xOccCode = getOccCode(xJob)
'End If
If xOccCode & "" = "" Then Exit Sub

Next_Step:

Select Case UptType
Case "A"
    If Not ifExistVadimOccCode(xOccCode) Then
        JobMasterBatchID = AddBatchVadim("A", UptDate)
        
        'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
        'City of Kawartha Lakes - Company Code
        'If glbCompSerial = "S/N - 2363W" Then
        '    If GetJobData(xOccCode, "JB_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
        '    Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:COMP_CODE", Null, xCompNo)
        'End If
        
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_CODE", Null, xOccCode)
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_NAME", Null, newDesc)
        
        'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
        'Ticket #23795: Town of Lasalle - New Salary Amount field, update the Rate Based On and Salary Amount accordingly
        'New Field added by Vadim - Ticket #14332, Ticket #14294
        If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
            If xGridType = "A" Then
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "S")
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:SALARY_AMOUNT", Null, xAnnSal)
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, 0)
            Else
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "H")
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
            End If
        Else
            Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "H")
            'Ticket #30112 - Town of Greater Napanee - Do not transfer the Hourly Rate/Base Rate
            If glbCompSerial <> "S/N - 2447W" Then
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
            End If
        End If
        
        'City of Timmins - Transfer Base Rate
        'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
        'If glbCompSerial = "S/N - 2375W" Then
        'Ticket #23795: Town of Lasalle - Update Salary Amount field for Annual Salary
        'If xGridType = "A" Then
        '    Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:SALARY_AMOUNT", Null, xAnnSal)
        'Else
        '    Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
        'End If
        
        Call CloseBatchVadim(JobMasterBatchID)
    Else
        JobMasterBatchID = AddBatchVadim("M", UptDate)
        
        'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
        'City of Kawartha Lakes - Company Code
        'If glbCompSerial = "S/N - 2363W" Then
        '    If GetJobData(xOccCode, "JB_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
        '    Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:COMP_CODE", Null, xCompNo)
        'End If
        
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_NAME", oldDesc, newDesc)
        
        'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
        'City of Timmins - Transfer Base Rate
        'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
        'If glbCompSerial = "S/N - 2375W" Then
        'Ticket #23795: Town of Lasalle - Update Salary Amount field for Annual Salary
        If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
            If xGridType = "A" Then
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "S")
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:SALARY_AMOUNT", Null, xAnnSal)
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, 0)
            Else
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "H")
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
            End If
        Else
            'Ticket #30112 - Town of Greater Napanee - Do not transfer the Hourly Rate/Base Rate
            If glbCompSerial <> "S/N - 2447W" Then
                Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
            End If
        End If
        
        Call CloseBatchVadim(JobMasterBatchID)
    End If
Case "M"
    VDClause = "OCC_CODE='" & xOccCode & "'"
    JobMasterBatchID = AddBatchVadim("M", UptDate)
    
    'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
    'City of Kawartha Lakes - Company Code
    'If glbCompSerial = "S/N - 2363W" Then
    '    If GetJobData(xOccCode, "JB_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
    '    Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:COMP_CODE", Null, xCompNo)
    'End If
    
    Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_NAME", oldDesc, newDesc, VDClause)
    
    'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
    'City of Timmins - Transfer Base Rate
    'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
    'If glbCompSerial = "S/N - 2375W" Then
    'Ticket #23795: Town of Lasalle - Update Salary Amount field for Annual Salary
    If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
        If xGridType = "A" Then
            Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "S", VDClause)
            Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:SALARY_AMOUNT", Null, xAnnSal, VDClause)
            Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, 0, VDClause)
        Else
            Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "H", VDClause)
            Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate, VDClause)
        End If
    Else
        'Ticket #30112 - Town of Greater Napanee - Do not transfer the Hourly Rate/Base Rate
        If glbCompSerial <> "S/N - 2447W" Then
            Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate, VDClause)
        End If
    End If
    
    Call CloseBatchVadim(JobMasterBatchID)
Case "D"
    JobMasterBatchID = AddBatchVadim("D", UptDate)
    Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_CODE", Null, xOccCode)
    Call CloseBatchVadim(JobMasterBatchID)
End Select
End Sub

Sub Passing_Salary_Grid_Vadim(StepNbr, ByVal OldValue, ByVal NewValue, UptDate, xJob, Optional xGrid)
Dim X
Dim VDClause
Dim UptType
Dim PayCodeInfo As PayCodeInfoType
Dim SalaryGirdBatchID
Dim xOccCode
Dim xCompNo As String
Dim xBaseRate
Dim xSalCD As String

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(SalaryGirdMaster) Then Exit Sub

If IsNull(OldValue) Then OldValue = 0
If IsNull(NewValue) Then NewValue = 0
OldValue = Val(OldValue)
NewValue = Val(NewValue)

'Ticket #23795: Town of Lasalle - They have Salary grid as well, so will need to update the Base Rate Based On and Salary field accordingly
'City of Timmins - Transfer Base Rate
'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
'If glbCompSerial = "S/N - 2375W" Then
    xBaseRate = GetJobData(xJob, "JB_S1")
    xSalCD = GetJobData(xJob, "JB_SALCD")   'Ticket #23795: Town of Lasalle
    If xBaseRate = "" Then
        xBaseRate = frmMPOSITIONS.medPayScale(1).Text
        If StepNbr = 1 Then     'Ticket #16090
            'Assign new rate
            xBaseRate = NewValue
        End If
        
        'Ticket #23795: Town of Lasalle
        If xSalCD = "" Then
            xSalCD = frmMPOSITIONS.lblSalCode.Caption
        End If
    Else
        If StepNbr = 1 And xBaseRate <> NewValue Then   'Ticket #16090
            'Assign new rate
            xBaseRate = NewValue
        End If
    End If
    
    'Get the Hourly Rate for the Base Rate as per their special calculation
    If glbCompSerial = "S/N - 2276W" Then
        xSalCD = "H"    'So their logic works to update with Hourly Rate only. Below we have added new field Salary Amount to be updated as well
        xBaseRate = Get_Hourly_Rate_NiagaraFalls(xJob)
        NewValue = Get_Hourly_Rate_NiagaraFalls(xJob, NewValue)
        If StepNbr = 1 Then 'Ticket #16090
            'Assign new rate
            xBaseRate = NewValue
        End If
    End If
'End If

If xBaseRate = "" Or IsNull(xBaseRate) Then xBaseRate = 0

'City of Timmins - Transfer Base Rate
If glbCompSerial = "S/N - 2375W" Then
    'do nothing
Else
    If OldValue = 0 And NewValue = 0 Then Exit Sub
End If

UptType = "M"
If OldValue = 0 Then UptType = "A"
If NewValue = 0 Then UptType = "D"

If glbLambton Then
    xOccCode = Left(xGrid, 1) & xJob & Mid(xGrid, 2)
Else
    xOccCode = xJob 'getOccCode(xJob)
End If
If xOccCode & "" = "" Then Exit Sub

Select Case UptType
Case "A"
    If Not ifExistVadimOccCode(xOccCode) Then
        SalaryGirdBatchID = AddBatchVadim(UptType, UptDate)
        
        'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
        'City of Kawartha Lakes - Company Code
        'If glbCompSerial = "S/N - 2363W" Then
        '    If GetJobData(xOccCode, "JB_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
        '    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:COMP_CODE", Null, xCompNo)
        'End If
        
        Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_CODE", Null, xOccCode)
        Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_NAME", Null, getIHRJobDesc(xOccCode))
        
        'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
        'Ticket #23795: Town of Lasalle
        'New Field added by Vadim - Ticket #14332, Ticket #14294
        If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, IIf(xSalCD = "A", "S", "H"))
        Else
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "H")
        End If
        
        'City of Timmins - Transfer Base Rate
        'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
        'If glbCompSerial = "S/N - 2375W" Then
        If StepNbr = 1 Then     'Ticket #16090
            'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
            'Ticket #23795: Town of Lasalle - Put the Salary Amount in Salary field
            If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
                If xSalCD = "A" Then
                    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:SALARY_AMOUNT", Null, xBaseRate)
                    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, 0)
                Else
                    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
                End If
            Else
                'Ticket #30112 - Town of Greater Napanee - Do not transfer the Hourly Rate/Base Rate
                If glbCompSerial <> "S/N - 2447W" Then
                    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
                End If
            End If
        End If
        'Else
        '    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, 0)
        'End If
        
        Call CloseBatchVadim(SalaryGirdBatchID)
    End If
    
    SalaryGirdBatchID = AddBatchVadim(UptType, UptDate)
    
    'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
    'City of Kawartha Lakes - Company Code
    'If glbCompSerial = "S/N - 2363W" Then
    '    If GetJobData(xOccCode, "JB_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
    '    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:COMP_CODE", Null, xCompNo)
    'End If
    
    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_CODE", Null, xOccCode)
    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_LEVEL", Null, StepNbr)
    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_LEVEL_NAME", Null, "Grid Step #" & StepNbr)
    
    'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
    'Ticket #23795: Town of Lasalle - Put the Salary Amount in Salary field
    If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
        If xSalCD = "A" Then
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_SALARY", Null, NewValue)
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_RATE", Null, 0)
        Else
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_RATE", Null, NewValue)
        End If
    Else
        Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_RATE", Null, NewValue)
    End If
    
    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_START_HOUR", Null, 0)
    
    'Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_SALARY", NULL NewValue)
    Call CloseBatchVadim(SalaryGirdBatchID)

    'Ticket #30364 - If the Step #1 is 0 to start for a NOT new Position then the Occupation Rate Level in Vadim was not getting updated. The routine below will do that.
    'City of Timmins - Update the Base Rate as well in Occupation table
    'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
    'If glbCompSerial = "S/N - 2375W" Then
    If StepNbr = 1 Then     'Ticket #16090
        VDClause = "OCC_CODE='" & xOccCode & "'"
        SalaryGirdBatchID = AddBatchVadim("M", UptDate)
        
        'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
        'Ticket #23795: Town of Lasalle
        If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, IIf(xSalCD = "A", "S", "H"), VDClause)
        End If
        
        'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
        'Ticket #23795: Town of Lasalle - Put the Salary Amount in Salary field
        If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
            If xSalCD = "A" Then
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:SALARY_AMOUNT", Null, xBaseRate, VDClause)
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, 0, VDClause)
            Else
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate, VDClause)
            End If
        Else
            'Ticket #30112 - Town of Greater Napanee - Do not transfer the Hourly Rate/Base Rate
            If glbCompSerial <> "S/N - 2447W" Then
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate, VDClause)
            End If
        End If
        
        Call CloseBatchVadim(SalaryGirdBatchID)
    End If
    'End If

Case "M"
    VDClause = "OCC_LEVEL=" & StepNbr
    SalaryGirdBatchID = AddBatchVadim("M", UptDate)
    
    'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
    'City of Kawartha Lakes - Company Code
    'If glbCompSerial = "S/N - 2363W" Then
    '    If GetJobData(xOccCode, "JB_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
    '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
    '    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:COMP_CODE", Null, xCompNo, VDClause)
    'End If
    
    'Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_LEVEL", StepNbr, StepNbr, VDClause)
    
    'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
    'Ticket #23795: Town of Lasalle - Put the Salary Amount in Salary field
    If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
        If xSalCD = "A" Then
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_SALARY", OldValue, NewValue, VDClause)
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_RATE", OldValue, 0, VDClause)
        Else
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_RATE", OldValue, NewValue, VDClause)
        End If
    Else
        Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_RATE", OldValue, NewValue, VDClause)
    End If
    
    Call CloseBatchVadim(SalaryGirdBatchID)
    
    'City of Timmins - Update the Base Rate as well in Occupation table
    'Hemu - Ticket #14294 - Jerry asked to pass Base Rate for all Vadim Clients.
    'If glbCompSerial = "S/N - 2375W" Then
    If StepNbr = 1 Then     'Ticket #16090
        VDClause = "OCC_CODE='" & xOccCode & "'"
        SalaryGirdBatchID = AddBatchVadim("M", UptDate)
        
        'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
        'Ticket #23795: Town of Lasalle
        If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
            Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, IIf(xSalCD = "A", "S", "H"), VDClause)
        End If
        
        'Ticket #25469 - City of Campbell River - Want Salary Amount transferred as well
        'Ticket #23795: Town of Lasalle - Put the Salary Amount in Salary field
        If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
            If xSalCD = "A" Then
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:SALARY_AMOUNT", Null, xBaseRate, VDClause)
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, 0, VDClause)
            Else
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate, VDClause)
            End If
        Else
            'Ticket #30112 - Town of Greater Napanee - Do not transfer the Hourly Rate/Base Rate
            If glbCompSerial <> "S/N - 2447W" Then
                Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate, VDClause)
            End If
        End If
        
        Call CloseBatchVadim(SalaryGirdBatchID)
    End If
    'End If
    
Case "D"
    VDClause = "OCC_LEVEL=" & StepNbr
    
    'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
    'City of Kawartha Lakes - Company Code
    'If glbCompSerial = "S/N - 2363W" Then
    '    If GetJobData(xOccCode, "JB_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
    '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
    'End If
    
    SalaryGirdBatchID = AddBatchVadim("D", UptDate)
    Call VadimInterface(SalaryGirdBatchID, xOccCode, "OCCRATE:OCC_LEVEL", StepNbr, Null, VDClause)
    Call CloseBatchVadim(SalaryGirdBatchID)
End Select

End Sub

Sub Passing_PAYINFO_Vadim(xPayID, HRField, ByVal OldValue, ByVal NewValue)
Dim X
Dim VDClause
Dim UptType
Dim PayCodeInfo As PayCodeInfoType
Dim xIHRType
Dim PAYINFOBatchID
Dim Y
Dim xCompNo, xUnion As String
Dim xUnionCode As String

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(Demographices) Then Exit Sub

xUnionCode = ""

If Right(HRField, 4) = "_ORG" Then
    If OldValue = NewValue Then Exit Sub
    
    If Len(OldValue) = 0 And Len(NewValue) <> 0 Then
        UptType = "A"
    ElseIf Len(OldValue) <> 0 And Len(NewValue) = 0 Then
        UptType = "D"
    Else
        'City of Timmins or City of Niagara Falls
        If glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W" Then
            UptType = "M"
            
        'Town of Aurora
        ElseIf glbCompSerial = "S/N - 2378W" Then
            If (OldValue = "0" Or OldValue = "") And Len(NewValue) <> 0 Then
                UptType = "A"
            ElseIf (OldValue <> "0" And OldValue <> "") And NewValue = "0" Then
                UptType = "D"
            End If
        
        'Dist. of Muskoka
        'WO# 129311 - CUPE Code - Create CUPE Union Due in iCity for Union Codes 1181, 1810 and PPT
        ElseIf (glbCompSerial = "S/N - 2373W") And _
               (NewValue = "1810" Or NewValue = "1811" Or NewValue = "PPT" Or _
               OldValue = "1810" Or OldValue = "1811" Or OldValue = "PPT") Then
               
            UptType = "M"
        Else
            Exit Sub
        End If
    End If
    
    'Ticket #25412 - Town of Greater Napanee
    If glbCompSerial = "S/N - 2447W" Then
        If (NewValue = "1" Or NewValue = "2") Then
            'New Value is Union Code
            If (OldValue <> "1" And OldValue <> "2") Then
                'Old value as non Union so add the Union Due pay code
                UptType = "A"
            ElseIf (OldValue = "1" Or OldValue = "2") Then
                'Old value was already a Union code so the Union Due pay code already exists, modify existing record.
                UptType = "M"
            End If
        Else
            'New Value is non Union
            If (OldValue <> "1" And OldValue <> "2") Then
                'Old value as non Union so do not do anything, there is no Union Due pay code
                Exit Sub
            ElseIf (OldValue = "1" Or OldValue = "2") Then
                'Old value was Union code so the Union Due pay code exists, delete the existing Union Due Pay Code as new codes are not Union codes.
                UptType = "D"
            End If
        End If
    End If
    
    'Dist. of Muskoka
    'WO# 129311 - CUPE Code - Create CUPE Union Due in iCity for Union Codes 1181, 1810 and PPT
    If (glbCompSerial = "S/N - 2373W") Then
        If (NewValue = "1810" Or NewValue = "1811" Or NewValue = "PPT" Or _
            OldValue = "1810" Or OldValue = "1811" Or OldValue = "PPT") Then
            If (NewValue = "1810" Or NewValue = "1811" Or NewValue = "PPT") And (OldValue <> "1810" And OldValue <> "1811" And OldValue <> "PPT") Then
                'Old Non Union Due value changed to Union Due value - Add Union Due Pay Code
                UptType = "A"
            ElseIf (OldValue = "1810" Or OldValue = "1811" Or OldValue = "PPT") And ((NewValue <> "1810" And NewValue <> "1811" And NewValue <> "PPT")) Then
                'Old Union Due value changed to Non Union Due value - Do not Delete Union Pay Code - it should never be deleted once it's added
                'UptType = "D"
                Exit Sub
            Else
                'Changing from one Union due to another - no change to Employee Trans Code Maintenance in iCity
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    
    If glbCompSerial = "S/N - 2375W" Then 'City of Timmins
        If NewValue = "P" Then
            If UptType = "M" Then
                NewValue = GetEmpData_PayrollID(xPayID, "ED_VADIM2") 'Rate Level
            Else
                NewValue = frmEESTATS.txtVadim1.Text
            End If
            PayCodeInfo.PayCode = "UNI"
            xUnionCode = "P"
        Else
            'Hard Codings for Rate Levels for Union Codes
            Select Case NewValue
                Case "1"
                    NewValue = 1
                Case "A"
                    NewValue = 2
                Case "C"
                    NewValue = 3
                Case "D"
                    NewValue = 4
                Case "E"
                    NewValue = 5
                Case "F"
                    NewValue = 6
                Case "G"
                    NewValue = 7
                Case "H"
                    NewValue = 8
                Case "I"
                    NewValue = 9
                Case "J"
                    NewValue = 10
                Case "K"
                    NewValue = 11
                'Case "P"
                '    NewValue = 15
            End Select
        End If
        
    'City of Kawartha Lakes
    ElseIf glbCompSerial = "S/N - 2363W" Then
        NewValue = GetEmpData(GetEmpData_PayrollID(xPayID, "ED_EMPNBR"), "ED_VADIM1") 'Rate Level
        
    'City of Niagara Falls
    ElseIf glbCompSerial = "S/N - 2276W" Then
        'Get the Rate Level from the mapped field
        If UptType = "M" Then
            NewValue = GetEmpData_PayrollID(xPayID, "ED_VADIM2") 'Rate Level
        ElseIf UptType = "D" And OldValue = "1" Then
            PayCodeInfo.PayCode = "UDI"
            xUnionCode = "1"
        Else
            If (NewHireForms.count > 0) And NewValue = "1" Then   'CUPE and New Hire
                PayCodeInfo.PayCode = "UDI"
                xUnionCode = "1"
            End If
            NewValue = frmEESTATS.txtVadim1.Text
        End If
        If NewValue = "" Then NewValue = "Null"
    Else
        'Town of Lasalle
        If glbCompSerial = "S/N - 2379W" Then
            If Not IsNumeric(NewValue) Then
                UptType = "D"
            End If
        End If
        
        OldValue = 0
        NewValue = 0
    End If
ElseIf HRField = "ED_CPP" Then
    If OldValue = NewValue Then Exit Sub
    If IsNull(OldValue) Then OldValue = "N"
    If OldValue = "" Then OldValue = "N"
    If NewValue <> "N" And OldValue = "N" Then UptType = "A"
    If NewValue = "N" And OldValue <> "N" Then UptType = "D"
    OldValue = 0
    NewValue = 0
    
    If glbCompSerial = "S/N - 2375W" Then 'City of Timmins
        OldValue = 0
        NewValue = "Null"
    End If
ElseIf HRField = "ED_OMERS" Or (glbCompSerial = "S/N - 2379W" And HRField = "ED_OMERS_1") Then
    If OldValue = NewValue Then Exit Sub
    'Town of Lasalle
    If glbCompSerial = "S/N - 2379W" And HRField = "ED_OMERS_1" Then
        'If oldValue = "" And (Val(NewValue) = "5" Or Val(NewValue) = "6") Then UptType = "A"
        If (OldValue = "" Or OldValue = "E") And (Val(NewValue) = "1" Or Val(NewValue) = "2") Then UptType = "A"
        'If Val(oldValue) = 1 And (Val(NewValue) <> "5" And Val(NewValue) <> "6") Then UptType = "D"
        If (Val(OldValue) = 1 Or Val(OldValue) = 2) And (Val(NewValue) <> "1" And Val(NewValue) <> "2") Then UptType = "D"
    Else
        If IsDate(NewValue) And Not IsDate(OldValue) Then UptType = "A"
        If Not IsDate(NewValue) And IsDate(OldValue) Then UptType = "D"
    
        OldValue = 0
        NewValue = 0
    End If
    
    'City of Timmins or City of Niagara Falls
    If glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2276W" Then
        OldValue = 0
        NewValue = "Null"
    End If
ElseIf HRField = "ED_UIC" Then
    'Town of Aurora - Ticket #20931
    UptType = "M"
    If OldValue = NewValue Then Exit Sub
    If glbCompSerial = "S/N - 2379W" Then   'Ticket #23795 - Town of LaSalle
        If OldValue = "" And Val(NewValue) <> "00" Then UptType = "A"
        If Val(OldValue) = 1 And NewValue = "00" Then UptType = "D"
    Else
        If OldValue = "" And Val(NewValue) = 1 Then UptType = "A"
        If Val(OldValue) = 1 And NewValue = "" Then UptType = "D"
    End If
ElseIf HRField = "ED_GROSSCD" Then
    'Town of Aurora - Ticket #20931
    'City of Kawartha Lakes - Ticket #21820
    UptType = "M"
    If OldValue = NewValue Then Exit Sub
    If IsNull(OldValue) Then OldValue = "N"
    If OldValue = "" Then OldValue = "N"
    If NewValue = "Y" And OldValue = "N" Then UptType = "A"
    If NewValue = "N" And OldValue <> "N" Then UptType = "D"
ElseIf HRField = "ED_WCBCODE" And (glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W") Then    'Town of Lasalle, Town of Greater Napanee - Ticket #24375
    'Town of Aurora
    UptType = "M"
    If OldValue = NewValue Then Exit Sub
    If OldValue = "" And NewValue <> "" Then UptType = "A"
    If NewValue = "" And OldValue <> "" Then UptType = "D"
ElseIf HRField = "ED_WCBCODE" And glbCompSerial = "S/N - 2458W" Then 'Ticket #25469 - City of Campbell River
    UptType = "M"
    If OldValue = NewValue Then Exit Sub
    If (OldValue = "" Or OldValue = "2") And NewValue = "1" Then UptType = "A"
    If NewValue = "2" And OldValue <> "" Then UptType = "D"
Else
    UptType = "M"
    If Val(OldValue) = 0 And Val(NewValue) = 0 Then Exit Sub
    If Val(OldValue) = 0 And Val(NewValue) <> 0 Then UptType = "A"
    
    'County of Lambton - do not delete the ITAX when Prov Amt is 0, Ticket #25412 - Town of Greater Napanee
    'Ticket #30250 - Campbell River - Do not delete Pay Code 800 for Income Tax
    If (glbCompSerial <> "S/N - 2355W") And (glbCompSerial <> "S/N - 2447W") And (glbCompSerial <> "S/N - 2458W") And (HRField = "ED_PROVAMT" And glbCompSerial <> "S/N - 2363W") Then
        If Val(OldValue) <> 0 And Val(NewValue) = 0 Then UptType = "D"
    Else
        If HRField = "ED_PROVAMT" Then
            If Val(OldValue) <> 0 And Val(NewValue) = 0 Then UptType = "M"
        Else
            If Val(OldValue) <> 0 And Val(NewValue) = 0 Then UptType = "D"
        End If
    End If
End If

'Ticket #24996 - City of Campbell River
'Town of Greater Napanee - Ticket #24375
'Town of Lasalle
If HRField = "ED_WCB" Or ((glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W") And HRField = "ED_UIC") Then
    xIHRType = "EI"
ElseIf HRField = "ED_WCBCODE" Then
    xIHRType = "WSIB"
ElseIf HRField = "ED_PROVAMT" Then
    xIHRType = "PROV. AMOUNT"
    If (glbCompSerial = "S/N - 2458W" Or glbCompSerial = "S/N - 2447W") Then 'Ticket #25469 - City of Campbell River, Ticket #25412 - Town of Greater Napanee
        xIHRType = "TAX"
    End If
ElseIf Right(HRField, 4) = "_ORG" Then
    xIHRType = "UNION"
ElseIf HRField = "ED_CPP" Then
    xIHRType = "CPP"
ElseIf HRField = "ED_OMERS" Or (glbCompSerial = "S/N - 2379W" And HRField = "ED_OMERS_1") Then
    xIHRType = "OMERS"
ElseIf HRField = "ED_GROSSCD" Then
    'Ticket #21820 - City of Kawartha Lakes
    If glbCompSerial = "S/N - 2363W" Then
        xIHRType = "PROV. AMOUNT"
    Else
        xIHRType = "TAX"
    End If
End If

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" And HRField = "ED_VADIM1" Then
    xIHRType = "UNION"
End If

'City of Timmins
If glbCompSerial = "S/N - 2375W" And HRField = "ED_VADIM2" Then
    xIHRType = "UNION"
    xUnionCode = GetEmpData_PayrollID(xPayID, "ED_ORG") 'Union
    If xUnionCode <> "P" Then Exit Sub
    
'City of Niagara Falls  - if Rate Level Change
ElseIf glbCompSerial = "S/N - 2276W" And HRField = "ED_VADIM2" Then
    xIHRType = "UNION"
    xUnionCode = GetEmpData_PayrollID(xPayID, "ED_ORG") 'Union
End If

Call getPayCodeInfo(PayCodeInfo, xIHRType)
If Len(PayCodeInfo.PayCode) = 0 Then Exit Sub
If Len(PayCodeInfo.PayFreq) = 0 Then Exit Sub

'City of Timmins - Ticket #12695
If glbCompSerial = "S/N - 2375W" Then
    If GetEmpData_PayrollID(xPayID, "ED_ORG") = "J" Then
        PayCodeInfo.PayFreq = "M"
    End If
End If

Y = 1

'Town of Aurora and Town of Lasalle has two OMERS pay code mapped (OMER and OME),
'Town of Greater Napanee - Ticket #24375,
'Ticket #24996 - City of Campbell River
If ((glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W") And HRField = "ED_OMERS") Or (glbCompSerial = "S/N - 2379W" And HRField = "ED_OMERS_1") Then
    Y = 2
    PayCodeInfo.PayCode = "OME"
End If

If glbCompSerial = "S/N - 2375W" Then 'City of Timmins
    If UptType = "M" And xUnionCode = "P" Then
        PayCodeInfo.PayCode = "UNN"
        UptType = "D"
    ElseIf xUnionCode = "P" Then
        PayCodeInfo.PayCode = "UNI"
        PayCodeInfo.PayFreq = "P"   'Ticket #16800
    ElseIf HRField = "ED_VADIM2" Then
        PayCodeInfo.PayCode = "UNI"
    ElseIf Right(HRField, 4) = "_ORG" And OldValue = "P" Then
        PayCodeInfo.PayCode = "UNI"
        UptType = "D"
    End If
    
'City of Niagara Falls
ElseIf glbCompSerial = "S/N - 2276W" Then
    If (NewHireForms.count > 0) And xUnionCode = "1" Then   'CUPE and New Hire
        PayCodeInfo.PayCode = "UDI"
    ElseIf UptType = "M" And xUnionCode = "CUPE" Then
        UptType = "D"
    ElseIf xUnionCode = "CUPE" Then
        PayCodeInfo.PayCode = "UDI"
    ElseIf HRField = "ED_VADIM2" Then
        PayCodeInfo.PayCode = "UDI"
    ElseIf Right(HRField, 4) = "_ORG" And OldValue = "1" Then
        PayCodeInfo.PayCode = "UDI"
        UptType = "D"
    End If
End If

For X = 1 To Y
    If X = 2 Then PayCodeInfo.PayCode = "OMER"
    
    Select Case UptType
    Case "A"
Timmins_Union:
        If isExistTransCode(xPayID, PayCodeInfo.PayCode) = 1 Then
            PAYINFOBatchID = AddBatchVadim("M")
            VDClause = "PAY_CODE='" & PayCodeInfo.PayCode & "'"
            
            'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
            'City of Kawartha Lakes - Company Code
            'If glbCompSerial = "S/N - 2363W" Then
            '    xUnion = GetEmpData(GetEmpData_PayrollID(xPayid, "ED_EMPNBR"), "ED_ORG")
            '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
            '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
            '    Call VadimInterface(PAYINFOBatchID, xPayid, "TRANSCODE:COMP_CODE", xCompNo, xCompNo, VDClause)
            'End If
            
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYCODE", PayCodeInfo.PayCode, PayCodeInfo.PayCode, VDClause)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:FREQCODE", PayCodeInfo.PayFreq, PayCodeInfo.PayFreq, VDClause)
            If HRField = "ED_PROVAMT" Then
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, Val(NewValue), VDClause)
            ElseIf HRField = "ED_GROSSCD" And NewValue = "Y" And glbCompSerial = "S/N - 2458W" Then
                'Ticket #25469 - City of Campbell River - Transfer Prov. Amount as well
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", GetEmpData_PayrollID(xPayID, "ED_PROVAMT"), GetEmpData_PayrollID(xPayID, "ED_PROVAMT"), VDClause)
            Else
                If glbCompSerial = "S/N - 2375W" And (HRField = "ED_CPP" Or HRField = "ED_OMERS") Then 'City of Timmins
                    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", OldValue, NewValue, VDClause)
                Else
                    'City of Niagara Falls - Do not pass 0, pass Null
                    If glbCompSerial = "S/N - 2276W" Then
                        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", 0, NewValue, VDClause)
                    ElseIf ((glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W") And (HRField = "ED_CPP" Or HRField = "ED_OMERS" Or Right(HRField, 4) = "_ORG")) Or (glbCompSerial = "S/N - 2379W" And HRField = "ED_OMERS_1") Then
                        'Ticket #24996 - City of Campbell River
                        'Town of Greater Napanee - Ticket #24375
                        'Town of Lasalle
                        'Town of Aurora - Ticket #20931 - As per mapping document, the Rate Level should be
                        'blank. But if blank Rate Level is passed then it does not update iCity so to avoid that
                        'we are not passing Rate Level.
                        'Dist. of Muskoka
                    ElseIf HRField = "ED_WCBCODE" And glbCompSerial = "S/N - 2458W" Then
                        'Ticket #25469 - City of Campbell River
                        'Do not pass Rate Level
                    Else
                        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", 0, Val(NewValue), VDClause)
                    End If
                End If
            End If
            Call CloseBatchVadim(PAYINFOBatchID)
        Else
            PAYINFOBatchID = AddBatchVadim("A")
            'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
            'City of Kawartha Lakes - Company Code
            'If glbCompSerial = "S/N - 2363W" Then
            '    xUnion = GetEmpData(GetEmpData_PayrollID(xPayid, "ED_EMPNBR"), "ED_ORG")
            '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
            '    Call VadimInterface(PAYINFOBatchID, xPayid, "TRANSCODE:COMP_CODE", xCompNo, xCompNo)
            'End If
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYROLL_ID", "", xPayID)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYCODE", "", PayCodeInfo.PayCode)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:FREQCODE", "", PayCodeInfo.PayFreq)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:REMAININGCALC", "", "9999")
            
            If HRField = "ED_PROVAMT" Then
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, Val(NewValue))
            ElseIf HRField = "ED_GROSSCD" And NewValue = "Y" And glbCompSerial = "S/N - 2458W" Then
                'Ticket #25469 - City of Campbell River - Transfer Prov. Amount as well
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, GetEmpData_PayrollID(xPayID, "ED_PROVAMT"))
            Else
                If glbCompSerial = "S/N - 2375W" And (HRField = "ED_CPP" Or HRField = "ED_OMERS") Then 'City of Timmins
                    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", OldValue, NewValue)
                Else
                    'City of Niagara Falls - Do not pass 0, pass Null
                    If glbCompSerial = "S/N - 2276W" Then
                        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", 0, NewValue)
                    ElseIf ((glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W") And (HRField = "ED_CPP" Or HRField = "ED_OMERS" Or Right(HRField, 4) = "_ORG")) Or (glbCompSerial = "S/N - 2379W" And HRField = "ED_OMERS_1") Then
                        'Ticket #24996 - City of Campbell River
                        'Town of Greater Napanee - Ticket #24375
                        'Town of Lasalle
                        'Town of Aurora - Ticket #20931 - As per mapping document, the Rate Level should be
                        'blank. But if blank Rate Level is passed then it does not update iCity so to avoid that
                        'we are not passing Rate Level.
                        'Dist. of Muskoka
                    ElseIf HRField = "ED_WCBCODE" And glbCompSerial = "S/N - 2458W" Then
                        'Ticket #25469 - City of Campbell River
                        'Do not pass Rate Level
                    Else
                        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", 0, Val(NewValue))
                    End If
                End If
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, 0)
            End If
            
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:EMPLR_MULTIPLIER", 0, 0)
            
            'City of Niagara Falls - Trans Date - if CUPE and New Hire -> DOH + 21
            'Ticket #12815 - getting error with this change
            'If glbCompSerial = "S/N - 2276W" Then
            '    If PayCodeInfo.PayCode = "UDI" Then
            '        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:DATE", "", DateAdd("d", 21, CVDate(frmEESTATS.dlpDate(7).Text)))
            '    End If
            'End If

            Call CloseBatchVadim(PAYINFOBatchID)
        End If
    Case "M"
        'City of Timmins - Because if employee is reinstated the Union is not created in Emp Trans Code Maint screen.
        If glbCompSerial = "S/N - 2375W" And Right(HRField, 4) = "_ORG" And isExistTransCode(xPayID, PayCodeInfo.PayCode) <> 1 Then
            PAYINFOBatchID = AddBatchVadim("A")
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYROLL_ID", "", xPayID)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYCODE", "", PayCodeInfo.PayCode)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:FREQCODE", "", PayCodeInfo.PayFreq)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:REMAININGCALC", "", "9999")
            If HRField = "ED_PROVAMT" Then
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, Val(NewValue))
            ElseIf HRField = "ED_GROSSCD" And NewValue = "Y" And glbCompSerial = "S/N - 2458W" Then
                'Ticket #25469 - City of Campbell River - Transfer Prov. Amount as well
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, GetEmpData_PayrollID(xPayID, "ED_PROVAMT"))
            Else
                'City of Timmins
                If glbCompSerial = "S/N - 2375W" And (HRField = "ED_CPP" Or HRField = "ED_OMERS") Then
                    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", OldValue, NewValue)
                Else
                    'City of Niagara Falls - Do not pass 0, pass Null
                    If glbCompSerial = "S/N - 2276W" Then
                        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", 0, NewValue)
                    Else
                        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", 0, Val(NewValue))
                    End If
                End If
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, 0)
            End If
            
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:EMPLR_MULTIPLIER", 0, 0)
            Call CloseBatchVadim(PAYINFOBatchID)
        Else
            PAYINFOBatchID = AddBatchVadim(UptType)
            VDClause = "PAY_CODE='" & PayCodeInfo.PayCode & "'"
            
            'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
            'City of Kawartha Lakes - Company Code
            'If glbCompSerial = "S/N - 2363W" Then
            '    xUnion = GetEmpData(GetEmpData_PayrollID(xPayid, "ED_EMPNBR"), "ED_ORG")
            '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
            '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
            '    Call VadimInterface(PAYINFOBatchID, xPayid, "TRANSCODE:COMP_CODE", xCompNo, xCompNo, VDClause)
            'End If
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYCODE", PayCodeInfo.PayCode, PayCodeInfo.PayCode, VDClause)
            Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:FREQCODE", PayCodeInfo.PayFreq, PayCodeInfo.PayFreq, VDClause)
        '    Call VadimInterface(PAYINFOBatchID, xPayid, "TRANSCODE:RATELEVEL", Val(oldValue), Val(NewValue), VDClause)
            If HRField = "ED_PROVAMT" Then
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", Val(OldValue), Val(NewValue), VDClause)
            ElseIf HRField = "ED_GROSSCD" And NewValue = "Y" And glbCompSerial = "S/N - 2458W" Then
                'Ticket #25469 - City of Campbell River - Transfer Prov. Amount as well
                Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", 0, GetEmpData_PayrollID(xPayID, "ED_PROVAMT"), VDClause)
            Else
                'City of Niagara Falls - Do not pass 0, pass Null
                If glbCompSerial = "S/N - 2276W" Then
                    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", Val(OldValue), NewValue, VDClause)
                ElseIf ((glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W") And (HRField = "ED_CPP" Or HRField = "ED_OMERS" Or Right(HRField, 4) = "_ORG")) Or (glbCompSerial = "S/N - 2379W" And HRField = "ED_OMERS_1") Then
                    'Ticket #24996 - City of Campbell River
                    'Town of Greater Napanee - Ticket #24375
                    'Town of Lasalle
                    'Town of Aurora - Ticket #20931 - As per mapping document, the Rate Level should be
                    'blank. But if blank Rate Level is passed then it does not update iCity so to avoid that
                    'we are not passing Rate Level.
                    'Dist. of Muskoka
                ElseIf HRField = "ED_WCBCODE" And glbCompSerial = "S/N - 2458W" Then
                    'Ticket #25469 - City of Campbell River
                    'Do not pass Rate Level
                Else
                    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", Val(OldValue), Val(NewValue), VDClause)
                End If
            End If
            Call CloseBatchVadim(PAYINFOBatchID)
        End If
    Case "D"
Delete_Union:
        PAYINFOBatchID = AddBatchVadim(UptType)
        VDClause = "PAY_CODE='" & PayCodeInfo.PayCode & "'"
        
        'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
        'City of Kawartha Lakes - Company Code
        'If glbCompSerial = "S/N - 2363W" Then
        '    xUnion = GetEmpData(GetEmpData_PayrollID(xPayid, "ED_EMPNBR"), "ED_ORG")
        '    If xUnion = "3" Then xCompNo = "2" Else xCompNo = "1"
        '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
        'End If
        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYCODE", PayCodeInfo.PayCode, "", VDClause)
        Call CloseBatchVadim(PAYINFOBatchID)
        
        'City of Timmins
        If glbCompSerial = "S/N - 2375W" Then
            If xUnionCode = "P" Then
               PayCodeInfo.PayCode = "UNI"
               PayCodeInfo.PayFreq = "P"   'Ticket #16800
               UptType = "A"
               GoTo Timmins_Union
            ElseIf OldValue = "P" And NewValue <> "" Then
               PayCodeInfo.PayCode = "UNN"
               UptType = "A"
               GoTo Timmins_Union
            End If
        
        'City of Niagara Falls
        ElseIf glbCompSerial = "S/N - 2276W" Then
            If OldValue = "1" And NewValue <> "" Then
               PayCodeInfo.PayCode = "UD"
               UptType = "A"
               GoTo Timmins_Union
            ElseIf OldValue = "1" And NewValue = "" Then
               PayCodeInfo.PayCode = "UD"
               UptType = "D"
               GoTo Delete_Union
            End If
        End If
    End Select
Next
End Sub

Sub getPayCodeInfo(PayCodeInfo As PayCodeInfoType, xIHRType, Optional xIHRCode)
Dim rsPayCode As New ADODB.Recordset
Dim SQLQ

SQLQ = "SELECT * FROM VADIM_PAYCODE WHERE IHR_TYPE='" & xIHRType & "'"
If Not IsMissing(xIHRCode) Then
    SQLQ = SQLQ & " AND IHR_CODE='" & xIHRCode & "'"
End If
rsPayCode.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If Not rsPayCode.EOF Then
    PayCodeInfo.PayCode = rsPayCode("PAY_CODE")
    PayCodeInfo.PayType = rsPayCode("PAY_TYPE_CODE")
    PayCodeInfo.PayTypeID = rsPayCode("PAY_TYPE_ID_CODE") & ""
    PayCodeInfo.PayFreq = Format(rsPayCode("PAY_CYCLE_FREQ_CODE"), "@")
End If
End Sub

Sub Passing_Bank_Vadim(Banks As Collection, xEmpnbr, Optional xPayID)
Dim X, xBNo
Dim PayIDs
Dim VDClause
Dim OldBank, NewBank
Dim OldBranch, NewBranch
Dim OldAccount, NewAccount
Dim OldType, NewType
Dim OldValue, NewValue
Dim UptType
Dim bankBatchID
Dim xCompNo As String

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

On Error GoTo Passing_Bank_Changes_Vadim_Err

If Not isTransfer(Banking) Then Exit Sub

PayIDs = Split(getPayrollIDs(xEmpnbr, xPayID), "|")

For X = 0 To UBound(PayIDs)
    For xBNo = 1 To 3
        'Ticket #24996 - City of Campbell River - They want to transfer Bank 2 and 3 - removed the Serial #
        'Town of Greater Napanee - Ticket #24375 - They allow more than 1 account (glbCompSerial = "S/N - 2447W")
        'Town of Lasalle
        'Town of Aurora - Ticket #20931 - as per mapping document - do not transfer Bank 2 and Bank 3
        If xBNo > 1 And (glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2379W") Then GoTo Next_PayID   'Or glbCompSerial = "S/N - 2458W"

        NewBank = Banks(xBNo)("BANK").NewValue
        OldBank = Banks(xBNo)("BANK").OldValue
        If NewBank <> "" Then NewBank = "0" & NewBank
        If OldBank <> "" Then OldBank = "0" & OldBank
        
        UptType = "M"
        If OldBank = "" And NewBank <> "" Then UptType = "A"
        If OldBank <> "" And NewBank = "" Then UptType = "D"
        
        If Not (OldBank = "" And NewBank = "") Then
            NewBranch = Banks(xBNo)("BRANCH").NewValue
            OldBranch = Banks(xBNo)("BRANCH").OldValue
            NewAccount = Banks(xBNo)("ACCOUNT").NewValue
            OldAccount = Banks(xBNo)("ACCOUNT").OldValue
            
            NewType = "": NewValue = "": OldType = "": OldValue = ""
            If NewBank <> "" Then
                If Val(Banks(xBNo)("AMTDEPOSIT").NewValue) <> 0 Then
                    NewType = "D": NewValue = Val(Banks(xBNo)("AMTDEPOSIT").NewValue)
                ElseIf Val(Banks(xBNo)("PCDEPOSIT").NewValue) <> 0 Then
                    NewType = "P": NewValue = Val(Banks(xBNo)("PCDEPOSIT").NewValue)
                Else
                    NewType = "R": NewValue = 0
                End If
            End If
            If OldBank <> "" Then
                If Val(Banks(xBNo)("AMTDEPOSIT").OldValue) <> 0 Then
                    OldType = "D"
                    OldValue = Val(Banks(xBNo)("AMTDEPOSIT").OldValue)
                ElseIf Val(Banks(xBNo)("PCDEPOSIT").OldValue) <> 0 Then
                    OldType = "P"
                    OldValue = Val(Banks(xBNo)("PCDEPOSIT").OldValue)
                Else
                    OldType = "R"
                    OldValue = 0
                End If
            End If
            
            VDClause = "BANK_CODE='" & OldBank & "'"
            VDClause = VDClause & " AND BANK_BRANCH_CODE='" & OldBranch & "'"
            VDClause = VDClause & " AND EMP_BANK_ACT='" & OldAccount & "'"
            
            'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
            'City of Kawartha Lakes - Company Code
            'If glbCompSerial = "S/N - 2363W" Then
            '    If GetEmpData(xEMPNBR, "ED_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
            '    VDClause = VDClause & " AND COMPANY_CODE= '" & xCompNo & "'"
            'End If
            
            If UptType = "A" Then
                bankBatchID = AddBatchVadim("A")
                
                'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
                'City of Kawartha Lakes - Company Code
                'If glbCompSerial = "S/N - 2363W" Then
                '    'If GetEmpData(xEmpnbr, "ED_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
                '    Call VadimInterface(bankBatchID, PayIDs(X), "BANK:COMP_CODE", Null, xCompNo)
                'End If
                
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:PAYROLL_ID", Null, PayIDs(X))
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:BANK", Null, NewBank)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:BRANCH", Null, NewBranch)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:ACCOUNT", Null, NewAccount)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:DISTTYPE", Null, NewType)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:DISTVALUE", Null, NewValue)
                Call CloseBatchVadim(bankBatchID)
            ElseIf UptType = "D" Then
                bankBatchID = AddBatchVadim("D")
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:PAYROLL_ID", PayIDs(X), Null, VDClause)
                Call CloseBatchVadim(bankBatchID)
            ElseIf OldBank <> NewBank Or OldBranch <> NewBranch Or OldAccount <> NewAccount Then
                bankBatchID = AddBatchVadim("D")
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:PAYROLL_ID", PayIDs(X), Null, VDClause)
                Call CloseBatchVadim(bankBatchID)
                
                bankBatchID = AddBatchVadim("A")
                
                'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
                'City of Kawartha Lakes - Company Code
                'If glbCompSerial = "S/N - 2363W" Then
                '    If GetEmpData(xEMPNBR, "ED_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
                '    Call VadimInterface(bankBatchID, PayIDs(X), "BANK:COMP_CODE", Null, xCompNo)
                'End If
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:PAYROLL_ID", Null, PayIDs(X))
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:BANK", Null, NewBank)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:BRANCH", Null, NewBranch)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:ACCOUNT", Null, NewAccount)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:DISTTYPE", Null, NewType)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:DISTVALUE", Null, NewValue)
                Call CloseBatchVadim(bankBatchID)
            Else
                bankBatchID = AddBatchVadim("M")
                
                'Hemu - Ticket #11246 - Remove Company Code passing to Vadim
                'City of Kawartha Lakes - Company Code
                'If glbCompSerial = "S/N - 2363W" Then
                '    If GetEmpData(xEMPNBR, "ED_ORG") = "3" Then xCompNo = "2" Else xCompNo = "1"
                '    Call VadimInterface(bankBatchID, PayIDs(X), "BANK:COMP_CODE", xCompNo, xCompNo, VDClause)
                'End If
                
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:DISTTYPE", OldType, NewType, VDClause)
                Call VadimInterface(bankBatchID, PayIDs(X), "BANK:DISTVALUE", OldValue, NewValue, VDClause)
                Call CloseBatchVadim(bankBatchID)
            End If
        End If
    Next xBNo
    
Next_PayID:
Next X

Exit Sub
Passing_Bank_Changes_Vadim_Err:
MsgBox Err.Description
Resume Next
End Sub

Public Function getVITField(VDEngField)
'Get from Vadim Import tables
Dim rsMap As New ADODB.Recordset

getVITField = ""

rsMap.Open "SELECT INFOHR_FIELD FROM VADIM_MAPPING WHERE USERSETUP<>0 AND VADIM_FIELD='" & VDEngField & "'", gdbAdoIhr001, adOpenForwardOnly
If Not rsMap.EOF Then
    getVITField = Format(rsMap("INFOHR_FIELD"), "@")
End If
rsMap.Close

End Function

Function getEmpValue(EMPField, xEmpnbr, xPayID)
'Get VALUE for a employee
Dim rsEmp As New ADODB.Recordset

getEmpValue = ""

If IsMissing(xPayID) Then xPayID = ""

If xPayID <> "" And EMPField = "ED_SECTION" Then
    EMPField = Replace(EMPField, "ED_", "JH_")
    rsEmp.Open "SELECT " & EMPField & " FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEmpnbr & " AND JH_PAYROLL_ID='" & xPayID & "'", gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        getEmpValue = Format(rsEmp(EMPField), "@")
    End If
    rsEmp.Close
Else
    rsEmp.Open "SELECT " & EMPField & " FROM HREMP WHERE ED_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        getEmpValue = Format(rsEmp(EMPField), "@")
    End If
    rsEmp.Close
End If
End Function

Sub PassDataToVadim(BatchID, VDTableField, VDEmpCode, OldValue, NewValue, Optional VDClause)
Dim rsFace As New ADODB.Recordset
Dim VDTemp, vdTable, VDField
Dim xAddressLine, xStreetNum
Dim xOAddressLine, xOStreetNum
Dim xPayType As String
Dim xVal
Dim xUpd As Boolean

VDTemp = Split(VDTableField, ".")
vdTable = VDTemp(0)
VDField = VDTemp(1)

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If glbCompSerial = "S/N - 2375W" And (IsNull(OldValue) Or IsNull(NewValue)) Then   'City of Timmins
    'let it be null
Else
    If IsNull(OldValue) Then OldValue = ""
    If IsNull(NewValue) Then NewValue = ""
End If

'And Not City of Niagara Falls
If (OldValue = "" And NewValue = "" And VDTableField = "EMPLOYEE.RPP_NUMBER" And glbCompSerial <> "S/N - 2276W") Then Exit Sub

'They don't want this logic - Nov 3rd 2014
'Ticket #24565 - District Municipality of Muskoka
'If glbCompSerial = "S/N - 2373W" Then
'    'Since we are transferring Employee # in place of Occupation Code/Job Group for Employee in Union '181W', the
'    'Occupation record needs to be created first with the Employee # as Occupation Code if not aleady existing.
'    If VDTableField = "EMPLOYEE.DEF_OCC_CODE" Or VDTableField = "EMPLOYEE.UNION_NUM" Then
'        'If Occupation Code or Union, and employee's Union is '181W' getting updated, check if the Occupation Code exists.
'        'If Code do not exists then add corresponding Occupation Code with Employee's Hourly Rate (Total)
'        If VDTableField = "EMPLOYEE.UNION_NUM" And UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
'            If frmEESTATS.clpCode(2) = "181W" Then
'                Call Passing_OccupationData_Vadim(VDEmpCode)
'            End If
'        ElseIf VDTableField = "EMPLOYEE.UNION_NUM" And UCase(MDIMain.ActiveForm.name) = "FRMEPOSITION" Then
'            If frmEPOSITION.clpCode(0) = "181W" Then
'                Call Passing_OccupationData_Vadim(VDEmpCode)
'            End If
'        Else
'            If GetEmpData_PayrollID(VDEmpCode, "ED_ORG") = "181W" Then
'                Call Passing_OccupationData_Vadim(VDEmpCode)
'            End If
'        End If
'    End If
'End If


rsFace.Open "SELECT * FROM SY_INTERFACE WHERE SY_BATCH_ID=-1", gdbPayroll, adOpenKeyset, adLockOptimistic
rsFace.AddNew
rsFace("SY_BATCH_ID") = BatchID
rsFace("DEST_TABLE") = vdTable
rsFace("DEST_KEY_VALUE") = VDEmpCode
rsFace("DEST_PROCESSED") = "N"
If Not IsMissing(VDClause) Then rsFace("DEST_WHERE_CLAUSE") = VDClause
rsFace("DEST_FIELD_NAME") = VDField

'Ticket #25469 - City of Campbell River
'Town of Lasalle - If Default Value set as D then pass as D
If glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2458W" Then
    If VDTableField = "EMPLOYEE.EMP_PAY_MODE_CODE" Then
        If OldValue = "Y" Or OldValue = "D" Then OldValue = "D" Else OldValue = "C"
        If NewValue = "Y" Or NewValue = "D" Then NewValue = "D" Else NewValue = "C"
    End If
Else
    If VDTableField = "EMPLOYEE.EMP_PAY_MODE_CODE" Then
        If OldValue = "Y" Then OldValue = "D" Else OldValue = "C"
        If NewValue = "Y" Then NewValue = "D" Else NewValue = "C"
    End If
End If

'Town of Lasalle - OMERS flag is not Date Dependent but Pension Code dependent.
If glbCompSerial <> "S/N - 2379W" Then
    If VDTableField = "EMPLOYEE.SUPERAN_APPLIC_FLAG" Then
        If IsDate(OldValue) Then OldValue = "Y" Else OldValue = "N"
        If IsDate(NewValue) Then NewValue = "Y" Else NewValue = "N"
    End If
End If

'If glbCompSerial = "S/N - 2375W" Then   ' City of Timmins
'    If VDTableField = "EMPLOYEE.RPP_NUMBER" Then
'        If IsDate(oldValue) Then
'            oldValue = "1"
'        ElseIf oldValue = "2" Then
'            oldValue = "2"
'        Else
'            oldValue = oldValue
'        End If
'        If IsDate(NewValue) Then
'            NewValue = "1"
'        ElseIf NewValue = "2" Then
'            NewValue = "2"
'        Else
'            NewValue = NewValue
'        End If
'    End If
'Else
' Not City of Kawartha Lakes & City of Timmins & City of Niagara Falls & Town of Lasalle
If (glbCompSerial <> "S/N - 2363W" And glbCompSerial <> "S/N - 2375W" And glbCompSerial <> "S/N - 2276W" And glbCompSerial <> "S/N - 2379W") Then
    If VDTableField = "EMPLOYEE.RPP_NUMBER" Then
        If IsDate(OldValue) Then OldValue = "1" Else OldValue = ""
        If IsDate(NewValue) Then NewValue = "1" Else NewValue = ""
    End If
End If
'Town of Lasalle - Only pass Numeric RPP #
If glbCompSerial = "S/N - 2379W" Then
    If VDTableField = "EMPLOYEE.RPP_NUMBER" Then
        If IsNumeric(OldValue) Then OldValue = OldValue Else OldValue = ""
        If IsNumeric(NewValue) Then NewValue = NewValue Else NewValue = ""
    End If
End If

If VDTableField = "EMPLOYEE.COUNCILLOR_FLAG" Then
    'Ticket #24996 - City of Campbell River
    If glbCompSerial = "S/N - 2458W" Then
        If OldValue = "MCL" Then OldValue = "Y" Else OldValue = "N"
        If NewValue = "MCL" Then NewValue = "Y" Else NewValue = "N"
        
    'Ticket #29959 - County of Lambton
    ElseIf glbCompSerial = "S/N - 2355W" Then
        If GetEmpData_PayrollID(VDEmpCode, "ED_DEPTNO") = "1000" And GetEmpData_PayrollID(VDEmpCode, "ED_DIV") = "COUN" Then
            If OldValue = "9" Then OldValue = "Y" Else OldValue = "N"
            NewValue = "Y"
        Else
            If OldValue = "9" Then OldValue = "Y" Else OldValue = "N"
            If NewValue = "9" Then NewValue = "Y" Else NewValue = "N"
        End If
    Else
        If OldValue = "9" Then OldValue = "Y" Else OldValue = "N"
        If NewValue = "9" Then NewValue = "Y" Else NewValue = "N"
    End If
End If
If VDTableField = "EMPLOYEE.EMP_OVERHEAD_PCT" Then
    If OldValue = "" Then OldValue = "0"
    If NewValue = "" Then NewValue = "0"
End If
If VDTableField = "EMPLOYEE.STAT_HOL_PCT" Then
    If OldValue = "" Then OldValue = "0"
    If NewValue = "" Then NewValue = "0"
End If
'Town of Lasalle - 00 = N and else Y
If glbCompSerial = "S/N - 2379W" And VDTableField = "EMPLOYEE.UIC_APPLIC_FLAG" Then
    If OldValue = "00" Then OldValue = "N" Else OldValue = "Y"
    If NewValue = "00" Then NewValue = "N" Else NewValue = "Y"
End If

'Town of Aurora - Ticket #20931 - do not pass phone extension as per mapping document
If glbCompSerial = "S/N - 2378W" Then
    If VDTableField = "CLIENT.PHONE2" Then
        If Len(NewValue) > 10 Then NewValue = Left(NewValue, 10)
    End If
End If

rsFace("DEST_OLD_VALUE") = OldValue
rsFace("DEST_NEW_VALUE") = NewValue

If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes
    'Do not transfer Occupation Code/Position Code if the Position Group <> VAD
    If VDTableField = "EMPLOYEE.DEF_OCC_CODE" Then
        If GetJobData(NewValue, "JB_GRPCD") <> "VAD" Then
            rsFace("DEST_NEW_VALUE") = ""
        End If
    End If
End If

'Ticket #24565 - District Municipality of Muskoka
If glbCompSerial = "S/N - 2373W" Then
    'Do not transfer Occupation Code/Position Code instead transfer Position Group Code of the Position.
    If VDTableField = "EMPLOYEE.DEF_OCC_CODE" Then
        'They do not want to this logic for 181W anymore - Nov 3rd 2014
        ''if Union = '181W' then transfer employee # to Occupation Code
        'If GetEmpData_PayrollID(VDEmpCode, "ED_ORG") = "181W" Then
        '    rsFace("DEST_OLD_VALUE") = VDEmpCode
        '    rsFace("DEST_NEW_VALUE") = VDEmpCode
        'Else
            rsFace("DEST_OLD_VALUE") = GetJobData(OldValue, "JB_GRPCD")
            rsFace("DEST_NEW_VALUE") = GetJobData(NewValue, "JB_GRPCD")
        'End If
    End If
End If

'Ticket #17450
If VDTableField = "OCCUPATION.OCC_NAME" Or VDTableField = "EMPLOYEE.EMP_JOB_TITLE" Then
    rsFace("DEST_NEW_VALUE") = Left(rsFace("DEST_NEW_VALUE"), 40)
End If

If glbCompSerial = "S/N - 2373W" Then 'District Municipality of Muskoka - Ticket #19113
    If VDTableField = "EMPLOYEE.DEPT_CODE" Then
        'Use Payroll Matrix to tranfer Dept Code
        rsFace("DEST_OLD_VALUE") = PayrollMatrix("DEPT", OldValue, "M_CONVERT1")
        rsFace("DEST_NEW_VALUE") = PayrollMatrix("DEPT", NewValue, "M_CONVERT1")
    End If
End If

rsFace.Update

'Ticket #24996 - Not for City of Campbell River - Changed the Job Title from Job Desc to Organization 1 Code Desc field.
If glbCompSerial <> "S/N - 2458W" Then
    If VDTableField = "EMPLOYEE.DEF_OCC_CODE" Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "EMP_JOB_TITLE"
        'Ticket #24565 - District Municipality of Muskoka
        If glbCompSerial = "S/N - 2373W" Then
            'Transfer the Position Group Code Description instead of Job Description
            rsFace("DEST_OLD_VALUE") = Left(GetTABLDesc("JBGC", (GetJobData(OldValue, "JB_GRPCD"))), 40)
            rsFace("DEST_NEW_VALUE") = Left(GetTABLDesc("JBGC", (GetJobData(NewValue, "JB_GRPCD"))), 40)
        Else
            rsFace("DEST_OLD_VALUE") = Left(getIHRJobDesc(OldValue), 40)
            rsFace("DEST_NEW_VALUE") = Left(getIHRJobDesc(NewValue), 40)
        End If
        rsFace.Update
    End If
End If

If glbCompSerial = "S/N - 2378W" Then 'Town of Aurora
    If VDTableField = "CLIENT.STREET1" Then
        If IsNumeric(Left(NewValue, 1)) And InStr(NewValue, " ") <> 0 Then
            xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xAddressLine
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "STREET_NUM"
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xStreetNum
            rsFace.Update
        End If
    End If
    If VDTableField = "CLIENT.STREET2" Then
        'They want Unit # to be updated with Street # as well.
        If IsNumeric(NewValue) Then
            'xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            'xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "UNIT_NUM"
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        End If
    End If
    If VDTableField = "EMPLOYEE.EMP_START_DATE" And OldValue = "" Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
        rsFace("DEST_OLD_VALUE") = OldValue
        rsFace("DEST_NEW_VALUE") = NewValue
        rsFace.Update
    End If
    
    'Ticket #20931 - As per the mapping documentation
    If VDTableField = "EMPLOYEE.HOURS_PER_DAY" Then
        'Only transfer if OMERS Date entered
        If IsDate(GetEmpData_PayrollID(VDEmpCode, "ED_OMERS")) Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
            rsFace("DEST_OLD_VALUE") = OldValue
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
            rsFace("DEST_OLD_VALUE") = OldValue
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        Else
            'Clear the values if no OMERS date; as per the documentation and testing by user.
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
            rsFace("DEST_OLD_VALUE") = OldValue
            rsFace("DEST_NEW_VALUE") = ""
            rsFace.Update
        
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
            rsFace("DEST_OLD_VALUE") = OldValue
            rsFace("DEST_NEW_VALUE") = ""
            rsFace.Update
        End If
    End If
    'Ticket #20931 - As per the mapping documentation and testing by the user. Not a new hire because it
    'duplicates these update fields.
    If VDTableField = "EMPLOYEE.SUPERAN_APPLIC_FLAG" And NewHireForms.count <= 0 Then
        If NewValue = "N" Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
            'rsFace("DEST_OLD_VALUE") = oldValue
            rsFace("DEST_NEW_VALUE") = ""
            rsFace.Update

            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
            'rsFace("DEST_OLD_VALUE") = oldValue
            rsFace("DEST_NEW_VALUE") = ""
            rsFace.Update
        ElseIf NewValue = "Y" Then
            xVal = GetJHData(frmEESTATS.lblEEID, "JH_DHRS", "")
            If Not IsNull(xVal) Or xVal = "" Then
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
                'rsFace("DEST_OLD_VALUE") = oldValue
                rsFace("DEST_NEW_VALUE") = xVal
                rsFace.Update
    
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
                'rsFace("DEST_OLD_VALUE") = oldValue
                rsFace("DEST_NEW_VALUE") = xVal
                rsFace.Update
            End If
        End If
    End If
    
ElseIf glbCompSerial = "S/N - 2375W" Then 'City of Timmins
    If VDTableField = "EMPLOYEE.NEXT_INCR_DATE" Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
        rsFace("DEST_OLD_VALUE") = OldValue
        rsFace("DEST_NEW_VALUE") = NewValue
        rsFace.Update
    End If
ElseIf glbCompSerial = "S/N - 2363W" Then 'City of Kawartha Lakes
    If VDTableField = "EMPLOYEE.PRIOR_TO_PROB_LVL" And NewValue = "" Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
        rsFace("DEST_OLD_VALUE") = OldValue
        rsFace("DEST_NEW_VALUE") = NewValue
        rsFace.Update
        
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "AFTER_PROBATION_LEVEL"
        rsFace("DEST_OLD_VALUE") = OldValue
        rsFace("DEST_NEW_VALUE") = NewValue
        rsFace.Update
        
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "DEF_OCC_CODE"
        rsFace("DEST_OLD_VALUE") = OldValue
        rsFace("DEST_NEW_VALUE") = NewValue
        rsFace.Update
        
        'rsFace.AddNew
        'rsFace("SY_BATCH_ID") = BatchID
        'rsFace("DEST_TABLE") = "EMPLOYEE"
        'rsFace("DEST_KEY_VALUE") = VDEmpCode
        'rsFace("DEST_FIELD_NAME") = "EMP_JOB_TITLE"
        'rsFace("DEST_OLD_VALUE") = oldValue
        'rsFace("DEST_NEW_VALUE") = NewValue
        'rsFace.Update
    End If
    If VDTableField = "EMPLOYEE.NEXT_INCR_DATE" Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
        rsFace("DEST_OLD_VALUE") = OldValue
        If GetJobData(frmESALARY.clpPostCode, "JB_GRPCD") <> "VAD" Then
            rsFace("DEST_NEW_VALUE") = ""
        Else
            xPayType = GetEmpData_PayrollID(VDEmpCode, "ED_LOC")
            If xPayType = "H" Then
                rsFace("DEST_NEW_VALUE") = IIf(Last_GridStep(VDEmpCode), "01/01/2099", NewValue)
            Else
                rsFace("DEST_NEW_VALUE") = NewValue
            End If
        End If
        rsFace.Update
    
        'If Next Review Date is populated then update Probation Level with current Step Level
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "PRIOR_TO_PROB_LVL"
        rsFace("DEST_OLD_VALUE") = ""
        If GetJobData(frmESALARY.clpPostCode, "JB_GRPCD") <> "VAD" Then
            rsFace("DEST_NEW_VALUE") = ""
        Else
            rsFace("DEST_NEW_VALUE") = Val(frmESALARY.lblSalaryGrade)
        End If
        rsFace.Update

        'Update After Probation Level with next Step Level. If the Step Level is the top most then
        'both Probation and After Probation Level should be same.
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "AFTER_PROBATION_LEVEL"
        rsFace("DEST_OLD_VALUE") = ""
        If GetJobData(frmESALARY.clpPostCode, "JB_GRPCD") <> "VAD" Then
            rsFace("DEST_NEW_VALUE") = ""
        Else
            If NewValue <> "" Then
                'xPayType = GetEmpData_PayrollID(VDEmpCode, "ED_LOC")
                'If xPayType = "H" Then
                    rsFace("DEST_NEW_VALUE") = IIf(Last_GridStep(VDEmpCode), Val(frmESALARY.lblSalaryGrade), Val(frmESALARY.lblSalaryGrade) + 1)
                'ElseIf (xPayType = "C") Or (xPayType = "P") Then
                '    rsFace("DEST_NEW_VALUE") = ""
                'End If
            Else
                rsFace("DEST_NEW_VALUE") = ""
            End If
        End If
        rsFace.Update
    End If
    
    If VDTableField = "CLIENT.STREET1" Then
        If IsNumeric(Left(NewValue, 1)) And InStr(NewValue, " ") <> 0 Then
            xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xAddressLine
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "STREET_NUM"
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xStreetNum
            rsFace.Update
        End If
    End If
ElseIf glbCompSerial = "S/N - 2276W" Then   'City of Niagara Falls - DOH = Probation Date
    'If VDTableField = "EMPLOYEE.EMP_START_DATE" Then   'Ticket #14285 - Do not pass anything to Probation Date
    '    rsFace.AddNew
    '    rsFace("SY_BATCH_ID") = BatchID
    '    rsFace("DEST_TABLE") = "EMPLOYEE"
    '    rsFace("DEST_KEY_VALUE") = VDEmpCode
    '    rsFace("DEST_PROCESSED") = "N"
    '    rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
    '    rsFace("DEST_OLD_VALUE") = oldValue
    '    rsFace("DEST_NEW_VALUE") = NewValue
    '    rsFace.Update
    If VDTableField = "EMPLOYEE.SICK_ACCR_HOURS" And NewHireForms.count <= 0 Then   'Default Sick Freq to "P"
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "SICK_ACCR_FREQ_CODE"
        rsFace("DEST_OLD_VALUE") = ""
        rsFace("DEST_NEW_VALUE") = "P"
        rsFace.Update
    ElseIf VDTableField = "CLIENT.STREET1" Then     'Split Address
        If IsNumeric(Left(NewValue, 1)) And InStr(NewValue, " ") <> 0 Then
            xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xAddressLine
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "STREET_NUM"
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xStreetNum
            rsFace.Update
        End If
    End If
ElseIf glbCompSerial = "S/N - 2373W" Then 'District Municipality of Muskoka - Ticket #19113
    If VDTableField = "CLIENT.STREET1" Then
        If IsNumeric(Left(NewValue, 1)) And InStr(NewValue, " ") <> 0 Then
            xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xAddressLine
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "STREET_NUM"
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xStreetNum
            rsFace.Update
        End If
    End If
    
    If VDTableField = "EMPLOYEE.HOURS_PER_DAY" And NewHireForms.count > 0 Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
        rsFace("DEST_OLD_VALUE") = 0
        rsFace("DEST_NEW_VALUE") = 0
        rsFace.Update
    
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "IN_LIEU_PCT"
        rsFace("DEST_OLD_VALUE") = 0
        rsFace("DEST_NEW_VALUE") = 0
        rsFace.Update
    End If
    
    'Ticket #24565 - They don't want Address 2 to be mapped to Unit Num anymore.
'    If VDTableField = "CLIENT.STREET2" Then
'        'They want Address 2 to be Numeric only so no need to extract Numeric value from Address 2 now
'        'They want Unit # to be updated with Street # as well.
'        If Len(NewValue) > 0 Then
'        'If IsNumeric(NewValue) Then
'            'xStreetNum = NewValue
'            xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
'
'            'xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
'            rsFace.AddNew
'            rsFace("SY_BATCH_ID") = BatchID
'            rsFace("DEST_TABLE") = "CLIENT"
'            rsFace("DEST_KEY_VALUE") = VDEmpCode
'            rsFace("DEST_PROCESSED") = "N"
'            rsFace("DEST_FIELD_NAME") = "UNIT_NUM"
'            rsFace("DEST_OLD_VALUE") = ""
'            rsFace("DEST_NEW_VALUE") = xStreetNum
'            rsFace.Update
'        End If
'    End If
    
    
    'They want to transfer for 181W as well now - Nov 3rd 2014
    'Ticket #24565 - if Union = '181W' then do not transfer Probation Date, Level and After Probation
    'If GetEmpData_PayrollID(VDEmpCode, "ED_ORG") <> "181W" Then
        'Ticket #24565 - Changed Probation Date from Next Increment Date to Last Increment Date.
        'Ticket #19113
        'If VDTableField = "EMPLOYEE.NEXT_INCR_DATE" Then
        If VDTableField = "EMPLOYEE.LAST_INCR_DATE" Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
            rsFace("DEST_OLD_VALUE") = OldValue
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        End If
    'End If
    
    'They don't want this logic for 181W employees as no Occupation Code creation with Employee # logic is needed. - Nov 3rd 2014
    'Ticket #24565 - Update Occupation Code with Employee # when Union changes for Employees in Union '181W'
    'If VDTableField = "EMPLOYEE.UNION_NUM" Then
    '    xUpd = False
    '    If VDTableField = "EMPLOYEE.UNION_NUM" And UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
    '        If frmEESTATS.clpCode(2) = "181W" Then
    '            xUpd = True
    '        End If
    '    ElseIf VDTableField = "EMPLOYEE.UNION_NUM" And UCase(MDIMain.ActiveForm.name) = "FRMEPOSITION" Then
    '        If frmEPOSITION.clpCode(0) = "181W" Then
    '            xUpd = True
    '        End If
    '    ElseIf GetEmpData_PayrollID(VDEmpCode, "ED_ORG") = "181W" Then
    '        xUpd = True
    '    End If
    '    If xUpd Then
    '        rsFace.AddNew
    '        rsFace("SY_BATCH_ID") = BatchID
    '        rsFace("DEST_TABLE") = "EMPLOYEE"
    '        rsFace("DEST_KEY_VALUE") = VDEmpCode
    '        rsFace("DEST_PROCESSED") = "N"
    '        rsFace("DEST_FIELD_NAME") = "DEF_OCC_CODE"
    '        rsFace("DEST_OLD_VALUE") = VDEmpCode
    '        rsFace("DEST_NEW_VALUE") = VDEmpCode
    '        rsFace.Update
    '    End If
    'End If
    
ElseIf glbCompSerial = "S/N - 2379W" Then   'Town of Lasalle - Transfer Pensionable Hours if Pension Code is entered
    If VDTableField = "EMPLOYEE.HOURS_PER_DAY" Then
        If Len(GetEmpData_PayrollID(VDEmpCode, "ED_PENSION")) > 0 Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
            rsFace("DEST_OLD_VALUE") = 0
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        End If
    End If
    
    If VDTableField = "EMPLOYEE.PRIOR_TO_PROB_LVL" And UCase(MDIMain.ActiveForm.name) = "FRMESALARY" Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
        'rsFace("DEST_OLD_VALUE") = OldValue
        rsFace("DEST_NEW_VALUE") = Format(frmESALARY.dlpDate(0), "yyyy/mm/dd")
        rsFace.Update
    End If
    If VDTableField = "EMPLOYEE.LAST_INCR_DATE" Then
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
        rsFace("DEST_OLD_VALUE") = OldValue
        rsFace("DEST_NEW_VALUE") = NewValue
        rsFace.Update
    End If
    If VDTableField = "EMPLOYEE.EMP_PAY_TYPE_CODE" Then
        'This is one is giving an error - it does not like blank values
        'Clear GL # if Payment Type is S otherwise by GL #
        'Ticket #24018 - Do not transfer or check for this when New Hire as it duplicates the fields in the
        'interface table.
        If NewValue <> "S" And NewHireForms.count <= 0 Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "DEF_ACCT_NUM"
            If NewValue = "S" Then
                rsFace("DEST_OLD_VALUE") = GetEmpData_PayrollID(VDEmpCode, "ED_GLNO")
                rsFace("DEST_NEW_VALUE") = ""
            Else
                rsFace("DEST_OLD_VALUE") = ""
                rsFace("DEST_NEW_VALUE") = GetEmpData_PayrollID(VDEmpCode, "ED_GLNO")
            End If
            rsFace.Update
        End If
        
        'Transfer Salary Distribution if Payment Type is S otherwise clear Salary Distribution
        rsFace.AddNew
        rsFace("SY_BATCH_ID") = BatchID
        rsFace("DEST_TABLE") = "EMPLOYEE"
        rsFace("DEST_KEY_VALUE") = VDEmpCode
        rsFace("DEST_PROCESSED") = "N"
        rsFace("DEST_FIELD_NAME") = "ACCT_DIST_CODE"
        rsFace("DEST_OLD_VALUE") = ""
        If NewValue = "S" Then
            rsFace("DEST_NEW_VALUE") = GetEmpData_PayrollID(VDEmpCode, "ED_SALDIST")
        Else
            rsFace("DEST_NEW_VALUE") = ""
        End If
        rsFace.Update
    End If
ElseIf glbCompSerial = "S/N - 2447W" Then 'Ticket #24375 - Town of Greater Napanee
    If VDTableField = "CLIENT.STREET1" Then
        If IsNumeric(Left(NewValue, 1)) And InStr(NewValue, " ") <> 0 Then
            xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xAddressLine
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "STREET_NUM"
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = xStreetNum
            rsFace.Update
        End If
    End If
    If VDTableField = "CLIENT.STREET2" Then
        If IsNumeric(NewValue) Then
            'xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            'xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "UNIT_NUM"
            rsFace("DEST_OLD_VALUE") = ""
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        End If
    End If
    If VDTableField = "EMPLOYEE.HOURS_PER_DAY" Then
        'Only transfer if OMERS Date entered
        If IsDate(GetEmpData_PayrollID(VDEmpCode, "ED_OMERS")) Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
            rsFace("DEST_OLD_VALUE") = OldValue
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
            rsFace("DEST_OLD_VALUE") = OldValue
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        Else
            'Does not like blank values on New Hire
            If (NewHireForms.count > 0) Then
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
                'rsFace("DEST_OLD_VALUE") = oldValue
                rsFace("DEST_NEW_VALUE") = 0
                rsFace.Update
            
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
                'rsFace("DEST_OLD_VALUE") = oldValue
                rsFace("DEST_NEW_VALUE") = 0
                rsFace.Update
            Else
                'Clear the values if no OMERS date; as per the documentation and testing by user.
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
                rsFace("DEST_OLD_VALUE") = OldValue
                rsFace("DEST_NEW_VALUE") = ""
                rsFace.Update
            
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
                rsFace("DEST_OLD_VALUE") = OldValue
                rsFace("DEST_NEW_VALUE") = ""
                rsFace.Update
            End If
        End If
    End If
    If VDTableField = "EMPLOYEE.SUPERAN_APPLIC_FLAG" And NewHireForms.count <= 0 Then
        If NewValue = "N" Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
            'rsFace("DEST_OLD_VALUE") = oldValue
            rsFace("DEST_NEW_VALUE") = ""
            rsFace.Update

            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
            'rsFace("DEST_OLD_VALUE") = oldValue
            rsFace("DEST_NEW_VALUE") = ""
            rsFace.Update
        ElseIf NewValue = "Y" Then
            'xVal = GetJHData(frmEESTATS.lblEEID, "JH_DHRS", "")
            xVal = GetJHData(GetEmpData_PayrollID(VDEmpCode, "ED_EMPNBR"), "JH_DHRS", "")
            If Not IsNull(xVal) Or xVal = "" Then
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
                'rsFace("DEST_OLD_VALUE") = oldValue
                rsFace("DEST_NEW_VALUE") = xVal
                rsFace.Update
    
                rsFace.AddNew
                rsFace("SY_BATCH_ID") = BatchID
                rsFace("DEST_TABLE") = "EMPLOYEE"
                rsFace("DEST_KEY_VALUE") = VDEmpCode
                rsFace("DEST_PROCESSED") = "N"
                rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
                'rsFace("DEST_OLD_VALUE") = oldValue
                rsFace("DEST_NEW_VALUE") = xVal
                rsFace.Update
            End If
        End If
    End If
    'If VDTableField = "EMPLOYEE.EMP_START_DATE" And OldValue = "" Then
    '    rsFace.AddNew
    '    rsFace("SY_BATCH_ID") = BatchID
    '    rsFace("DEST_TABLE") = "EMPLOYEE"
    '    rsFace("DEST_KEY_VALUE") = VDEmpCode
    '    rsFace("DEST_PROCESSED") = "N"
    '    rsFace("DEST_FIELD_NAME") = "PROBATION_DATE"
    '    rsFace("DEST_OLD_VALUE") = OldValue
    '    rsFace("DEST_NEW_VALUE") = NewValue
    '    rsFace.Update
    'End If
ElseIf glbCompSerial = "S/N - 2458W" Then 'Ticket #24996 - City of Campbell River
    If VDTableField = "CLIENT.STREET1" Then
        If IsNumeric(Left(NewValue, 1)) And InStr(NewValue, " ") <> 0 Then
            'Old values
            If Len(OldValue) > 0 Then
                xOStreetNum = Left(OldValue, InStr(OldValue, " ") - 1)
                xOAddressLine = Mid(OldValue, InStr(OldValue, " ") + 1)
            End If
            
            'New values
            xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            
            rsFace("DEST_OLD_VALUE") = xOAddressLine
            rsFace("DEST_NEW_VALUE") = xAddressLine
            
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "STREET_NUM"
            rsFace("DEST_OLD_VALUE") = xOStreetNum
            rsFace("DEST_NEW_VALUE") = xStreetNum
            rsFace.Update
        End If
    End If
    If VDTableField = "CLIENT.STREET2" Then
        If IsNumeric(NewValue) Then
            'xStreetNum = Left(NewValue, InStr(NewValue, " ") - 1)
            'xAddressLine = Mid(NewValue, InStr(NewValue, " ") + 1)
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "CLIENT"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "UNIT_NUM"
            If IsNumeric(OldValue) Then
                rsFace("DEST_OLD_VALUE") = OldValue
            Else
                rsFace("DEST_OLD_VALUE") = ""
            End If
            rsFace("DEST_NEW_VALUE") = NewValue
            rsFace.Update
        End If
    End If
    If VDTableField = "EMPLOYEE.HOURS_PER_DAY" Then
        'They are not in ONTARIO so no OMERS Date logic for them
        'Only transfer if OMERS Date entered
'        If IsDate(GetEmpData_PayrollID(VDEmpCode, "ED_OMERS")) Then
'            rsFace.AddNew
'            rsFace("SY_BATCH_ID") = BatchID
'            rsFace("DEST_TABLE") = "EMPLOYEE"
'            rsFace("DEST_KEY_VALUE") = VDEmpCode
'            rsFace("DEST_PROCESSED") = "N"
'            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
'            rsFace("DEST_OLD_VALUE") = oldValue
'            rsFace("DEST_NEW_VALUE") = NewValue
'            rsFace.Update
'
'            rsFace.AddNew
'            rsFace("SY_BATCH_ID") = BatchID
'            rsFace("DEST_TABLE") = "EMPLOYEE"
'            rsFace("DEST_KEY_VALUE") = VDEmpCode
'            rsFace("DEST_PROCESSED") = "N"
'            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
'            rsFace("DEST_OLD_VALUE") = oldValue
'            rsFace("DEST_NEW_VALUE") = NewValue
'            rsFace.Update
'        Else
'            'Clear the values if no OMERS date; as per the documentation and testing by user.
'            rsFace.AddNew
'            rsFace("SY_BATCH_ID") = BatchID
'            rsFace("DEST_TABLE") = "EMPLOYEE"
'            rsFace("DEST_KEY_VALUE") = VDEmpCode
'            rsFace("DEST_PROCESSED") = "N"
'            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
'            rsFace("DEST_OLD_VALUE") = oldValue
'            rsFace("DEST_NEW_VALUE") = ""
'            rsFace.Update
'
'            rsFace.AddNew
'            rsFace("SY_BATCH_ID") = BatchID
'            rsFace("DEST_TABLE") = "EMPLOYEE"
'            rsFace("DEST_KEY_VALUE") = VDEmpCode
'            rsFace("DEST_PROCESSED") = "N"
'            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
'            rsFace("DEST_OLD_VALUE") = oldValue
'            rsFace("DEST_NEW_VALUE") = ""
'            rsFace.Update
'        End If
        
        If UCase(MDIMain.ActiveForm.name) = "FRMEPOSITION" Then
            xVal = GetJobData(frmEPOSITION.clpJob, "JB_DHRS")
        Else
            xVal = GetJobData(GetJHData(GetEmpData_PayrollID(VDEmpCode, "ED_EMPNBR"), "JH_JOB", ""), "JB_DHRS")
        End If
        If Not IsNull(xVal) And xVal <> "" Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "EMP_RPP_HRS"
            'rsFace("DEST_OLD_VALUE") = oldValue
            rsFace("DEST_NEW_VALUE") = xVal
            rsFace.Update
        End If

        If (NewHireForms.count > 0) Then
            rsFace.AddNew
            rsFace("SY_BATCH_ID") = BatchID
            rsFace("DEST_TABLE") = "EMPLOYEE"
            rsFace("DEST_KEY_VALUE") = VDEmpCode
            rsFace("DEST_PROCESSED") = "N"
            rsFace("DEST_FIELD_NAME") = "ACCRUE_HRS_DAY"
            'rsFace("DEST_OLD_VALUE") = oldValue
            rsFace("DEST_NEW_VALUE") = 0
            rsFace.Update
        End If
    End If
End If
rsFace.Close

End Sub
'Sub Add_OccupationRecordToVadim(xOccCode)
Sub Passing_OccupationData_Vadim(xOccCode, Optional OldValue, Optional NewValue, Optional UDate)
Dim UptDate
Dim JobMasterBatchID
Dim xBaseRate
Dim newDesc
Dim empName
Dim VDClause

If Not IsMissing(UDate) Then
    UptDate = UDate
Else
    UptDate = Date
End If

    'Get Occupation Description as Employee Name
    empName = GetEmpData_PayrollID(xOccCode, "ED_FNAME") & " " & GetEmpData_PayrollID(xOccCode, "ED_SURNAME")


    If Not ifExistVadimOccCode(xOccCode) Then
        'Occupation Name is same for all Employee
        newDesc = "W&S PREMIUM"
        
        'Employee's Hourly Rate will be the Base Rate
        xBaseRate = GetSHData(GetEmpData_PayrollID(xOccCode, "ED_EMPNBR"), "SH_TOTAL", 0)
        
        JobMasterBatchID = AddBatchVadim("A", UptDate)
                
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_CODE", Null, xOccCode)
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_NAME", Null, newDesc)
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_DESC", Null, empName)
        
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:RATE_BASED_ON", Null, "H")
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", Null, xBaseRate)
        
        
        Call CloseBatchVadim(JobMasterBatchID)
    Else
        VDClause = "OCC_CODE='" & xOccCode & "'"
        JobMasterBatchID = AddBatchVadim("M", UptDate)
        
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_DESC", Null, empName, VDClause)
        Call VadimInterface(JobMasterBatchID, xOccCode, "OCCUPATION:OCC_LEVEL_RATE", OldValue, NewValue, VDClause)
        
        Call CloseBatchVadim(JobMasterBatchID)
    End If
End Sub

Function getIHRJobDesc(ByVal xJobCode)
Dim rsJB As New ADODB.Recordset
If glbLambton Then
    xJobCode = Mid(xJobCode, 2, 4)
End If
rsJB.Open "SELECT JB_DESCR FROM HRJOB WHERE JB_CODE='" & xJobCode & "'", gdbAdoIhr001, adOpenForwardOnly
If rsJB.EOF Then
    getIHRJobDesc = "Unknown"
Else
    getIHRJobDesc = rsJB("JB_DESCR")
End If
rsJB.Close
End Function

Function Next_Step_Available(PayID, NewValue)
Dim rsJobHis As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ As String

'Ticket #22682 - Release 8.0: Increased Salary Grid Steps from 11 to 15 -> 20
'If Val(NewValue) = 11 Then
'If Val(NewValue) = 15 Then
If Val(NewValue) = 20 Then
    Next_Step_Available = False
    Exit Function
End If
SQLQ = "SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_EMPNBR = (SELECT ED_EMPNBR FROM HREMP WHERE ED_PAYROLL_ID = '" & PayID & "') AND JH_CURRENT<>0"
rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsJobHis.EOF Then
    SQLQ = "SELECT JB_S" & Val(NewValue) + 1 & " FROM HRJOB WHERE JB_CODE = '" & rsJobHis("JH_JOB") & "'"
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsJOB.EOF Then
        If rsJOB("JB_S" & Val(NewValue) + 1) <> 0 And Not IsNull(rsJOB("JB_S" & Val(NewValue) + 1)) Then
            Next_Step_Available = True
        Else
            Next_Step_Available = False
        End If
    Else
        Next_Step_Available = False
    End If
    rsJOB.Close
Else
    Next_Step_Available = False
End If
rsJobHis.Close
End Function

Function Last_GridStep(PayID)
Dim rsSalHis As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ As String

'SQLQ = "SELECT SH_JOB,SH_GRADE FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = (SELECT ED_EMPNBR FROM HREMP WHERE ED_PAYROLL_ID = '" & PayID & "') AND SH_CURRENT<>0"
'rsSalHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'If Not rsSalHis.EOF Then
    'Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    'If Val(frmESALARY.lblSalaryGrade) = 11 Then
    'If Val(frmESALARY.lblSalaryGrade) = 15 Then
    If Val(frmESALARY.lblSalaryGrade) = 20 Then
        Last_GridStep = True
        Exit Function
    End If
    SQLQ = "SELECT JB_S" & Val(frmESALARY.lblSalaryGrade) + 1 & " FROM HRJOB WHERE JB_CODE = '" & frmESALARY.clpPostCode & "'"
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsJOB.EOF Then
        If rsJOB("JB_S" & Val(frmESALARY.lblSalaryGrade) + 1) <> 0 And Not IsNull(rsJOB("JB_S" & Val(frmESALARY.lblSalaryGrade) + 1)) Then
            Last_GridStep = False
        Else
            Last_GridStep = True
        End If
    Else
        Last_GridStep = True
    End If
    rsJOB.Close
'Else
'    Last_GridStep = True
'End If
'rsSalHis.Close
End Function

Sub AddVadimTables_Map()
Call AddVadimTables
Call AddMap
End Sub

Sub AddVadimTables()
Dim X

For X = 1 To VDTables.count
    VDTables.Remove 1
Next

Call AddFieldInfo(VDTables, "EMPLOYEE.COMPANY_CODE", "Company Code", aString, 2, , True, "COMP_CODE")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_NUM", "Employee Number", aString, 10, , True, "JH_PAYROLL_ID")
Call AddFieldInfo(VDTables, "EMPLOYEE.DEF_OCC_CODE", "Occupation Code", aString, 6, , False, "JH_JOB")
Call AddFieldInfo(VDTables, "EMPLOYEE.DEF_ACCT_NUM", "GL Number", aNumber, 30, 0, False, "JH_GLNO")
'Call AddFieldInfo(VDTables, "EMPLOYEE.DEF_EQUIP_CODE", "Equipment Code", aString, 4, , False, "VIT:Equipment")
Call AddFieldInfo(VDTables, "EMPLOYEE.DEPT_CODE", "Dept Code", aString, 4, , True, "JH_DEPTNO")
'Call AddFieldInfo(VDTables, "EMPLOYEE.CLIENT_CODE", "Client Code", aString, 10, , False)
Call AddFieldInfo(VDTables, "EMPLOYEE.UNION_NUM", "Union Number", aString, 1, , False, "JH_ORG")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_CLASS_CODE", "Employee Class Code", aString, 1, , False, "DFLT:EMP_CLASS_CODE:1")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_SIN", "Social Insurance Number", aString, 9, , False, "ED_SIN")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_BIRTH_DATE", "Employee Birth Date", aDate, 8, , False, "ED_DOB")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_CAT_CODE", "Employee Category", aString, 4, , True, "JH_PAYROLL_CATEGORY")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_PAY_TYPE_CODE", "Employee Pay Type Code", aString, 1, , True, "VIT:Payment Type")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_PAY_FREQ", "Employee Pay Frequency", aNumber, 5, 0, True, "SH_PAYP")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_START_DATE", "Employee Start date", aDate, 8, , True, "VIT:Start Date")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_UIC_START_DATE", "Employee UIC Start Date", aDate, 8, , False, "VIT:EI Start Date")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_TERM_DATE", "Employee Termination Date", aDate, 8, , False, "TERM:EMP_DATE")
Call AddFieldInfo(VDTables, "EMPLOYEE.NEXT_INCR_DATE", "Next Increment Date", aDate, 8, , False, "SH_NEXTDAT")
Call AddFieldInfo(VDTables, "EMPLOYEE.LAST_INCR_DATE", "Last Increment Date", aDate, 8, , False, "SH_EDATE")
Call AddFieldInfo(VDTables, "EMPLOYEE.SENIORITY_DATE", "Seniority Date", aDate, 8, , False, "VIT:Seniority Date")
Call AddFieldInfo(VDTables, "EMPLOYEE.MAX_BANK_TIME_HRS", "Maximum Bank Time Hours", aNumber, 12, 2, False, "ED_SUPCODE")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_TD1_AMT", "Employee TD1 Amount", aNumber, 12, 2, True, "ED_TD1DOL")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_TD3_AMT", "Employee TD3 Amount", aNumber, 12, 2, True, "ED_TD3")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_TD3_PERCENT", "Employee TD3 Percent", aNumber, 5, 4, False, "ED_TD3PC")
'Call AddFieldInfo(VDTables, "EMPLOYEE.NORTH_HOUS_ALLOW", "Northern Housing Allowance", aNumber, 12, 2, True)
Call AddFieldInfo(VDTables, "EMPLOYEE.CPP_APPLIC_FLAG", "CPP Applicable Flag", aString, 1, , True, "ED_CPP")
Call AddFieldInfo(VDTables, "EMPLOYEE.SUPERAN_APPLIC_FLAG", "Superanuation Applicable Flag", aString, 1, , True, "ED_OMERS")
'If glbCompSerial = "S/N - 2375W" Then  'City of Timmins
    'Call AddFieldInfo(VDTables, "EMPLOYEE.RPP_NUMBER", "RPP Number", aNumber, 5, 0, False, "ED_NORMALR")
    Call AddFieldInfo(VDTables, "EMPLOYEE.RPP_NUMBER", "RPP Number", aNumber, 5, 0, False, "ED_OMERS_1")
'Else
    'Call AddFieldInfo(VDTables, "EMPLOYEE.RPP_NUMBER", "RPP Number", aNumber, 5, 0, False, "ED_OMERS_1")
'End If
Call AddFieldInfo(VDTables, "EMPLOYEE.UIC_APPLIC_FLAG", "UIC Applicable Flag", aString, 1, , True, "ED_UIC")
Call AddFieldInfo(VDTables, "EMPLOYEE.TAX_APPLIC_FLAG", "Income Tax Applicable Flag", aString, 1, , True, "ED_GROSSCD")

Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_PAY_MODE_CODE", "Employee Pay Mode", aString, 1, , True, "ED_DDI")
Call AddFieldInfo(VDTables, "EMPLOYEE.HOURS_PER_DAY", "Hours per Day", aNumber, 12, 2, True, "JH_DHRS")
Call AddFieldInfo(VDTables, "EMPLOYEE.VAC_PAY_PCT", "Vacation Pay Percent", aNumber, 5, 4, True, "ED_VACPC")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_OVERHEAD_PCT", "Employee Overhead Percentage", aNumber, 5, 4, True, "ED_VADIM1")
'Call AddFieldInfo(VDTables, "EMPLOYEE.FIREMAN_POINTS", "Fireman Points", aNumber, 5, 0, False)
Call AddFieldInfo(VDTables, "EMPLOYEE.SICK_ACCR_HOURS", "Sick Accrual Hours", aNumber, 12, 2, True, "DFLT:SICK_ACCR_HOURS:0")
Call AddFieldInfo(VDTables, "EMPLOYEE.SICK_ACCR_FREQ_CODE", "Sick Accrual Freq Code", aString, 1, , False, "ED_VADIM2_SICK")
Call AddFieldInfo(VDTables, "EMPLOYEE.SICK_ACCR_MAXIMUM", "Sick Accrual Maximum", aNumber, 12, 2, True, "DFLT:SICK_ACCR_MAXIMUM:0")
Call AddFieldInfo(VDTables, "EMPLOYEE.VAC_ACC_HOURS", "Vacation Accrual Hours", aNumber, 12, 2, True, "DFLT:VAC_ACC_HOURS:0")
'Call AddFieldInfo(VDTables, "EMPLOYEE.VAC_ACC_FREQ_CODE", "Vacation Accrual Freq Code", aString, 1, , False)
Call AddFieldInfo(VDTables, "EMPLOYEE.VAC_ACC_MAX", "Vacation Accrual Maximum", aNumber, 12, 2, True, "DFLT:VAC_ACC_MAX:0")
Call AddFieldInfo(VDTables, "EMPLOYEE.PAY_VAC_PER_FLAG", "Pay Vacation Every Period Flag", aString, 1, , True, "ED_VACPC_1")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_DEFAULT_JOB", "Employee Class Group", aString, 10, , False, "JH_JOB_1")
Call AddFieldInfo(VDTables, "EMPLOYEE.STAT_HOL_PCT", "Statutory Holiday Percent", aNumber, 5, 4, True, "ED_VADIM2")
Call AddFieldInfo(VDTables, "EMPLOYEE.COUNCILLOR_FLAG", "Councillor Flag", aString, 1, , True, "RESET:ED_EMPTYPE")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_JOB_TITLE", "Employee Job Title", aString, 40, , False, "JB_DESCR")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_ACTIVE_FLAG", "Employee Active Flag", aString, 1, , True, "DFLT:EMP_ACTIVE_FLAG:Y")
'Call AddFieldInfo(VDTables, "EMPLOYEE.NON_REF_TAX_CRD", "Non Refundable Tax Credit", aNumber, 12, 2, True)
Call AddFieldInfo(VDTables, "EMPLOYEE.ANN_VAC_HRS", "Annual Vacation Hours", aNumber, 12, 2, True, "DFLT:ANN_VAC_HRS:0")
Call AddFieldInfo(VDTables, "EMPLOYEE.VOL_FIREMAN_FLAG", "Volunteer Fireman Flag", aString, 1, , True, "DFLT:VOL_FIREMAN_FLAG:N")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_HOURS", "Employee Hours", aNumber, 12, 2, True, "DFLT:EMP_HOURS:0")
Call AddFieldInfo(VDTables, "EMPLOYEE.PROBATION_DATE", "Probation Date", aDate, 8, , False, "DFLT:PROBATION_DATE:2099/1/1")
Call AddFieldInfo(VDTables, "EMPLOYEE.PRIOR_TO_PROB_LVL", "Prior to Probation Level", aNumber, 5, 0, False, "SH_GRADE")
Call AddFieldInfo(VDTables, "EMPLOYEE.AFTER_PROBATION_LEVEL", "After Probation Level", aNumber, 5, 0, False, "SH_GRADE_SAME")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_GENDER", "Employee Gender", aString, 1, , False, "ED_SEX")
'Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_RACE", "Employee Race", aString, 2, , False)
Call AddFieldInfo(VDTables, "EMPLOYEE.MARITAL_STATUS", "Marital Status", aString, 2, , False, "ED_MSTAT")
'Call AddFieldInfo(VDTables, "EMPLOYEE.SPOUSE_NAME", "Spouse Name", aString, 40, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.NEXT_OF_KIN", "Next of Kin", aString, 40, , False)
Call AddFieldInfo(VDTables, "EMPLOYEE.DRIVERS_LICENCE_INFO", "Drivers Licence Information", aString, 100, , False, "ED_DRIVERLIC")
'Call AddFieldInfo(VDTables, "EMPLOYEE.QUALIFICATIONS", "Qualifications", aString, 255, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.EXPERIENCE", "Experience", aString, 255, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.WCB_INFO", "WCB Information", aString, 255, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.DISABILITY", "Disability", aString, 40, , False)
Call AddFieldInfo(VDTables, "EMPLOYEE.ACCT_DIST_CODE", "Salary Distribution", aString, 6, , False, "ED_SALDIST")
'Call AddFieldInfo(VDTables, "EMPLOYEE.INS_USER", "Insert User", aString, 30, , True)
'Call AddFieldInfo(VDTables, "EMPLOYEE.INS_DATE", "Insert Date", aDate, 8, , True)
'Call AddFieldInfo(VDTables, "EMPLOYEE.UPD_USER", "Update User", aString, 30, , True)
'Call AddFieldInfo(VDTables, "EMPLOYEE.UPD_DATE", "Update Date", aDate, 8, , True)
Call AddFieldInfo(VDTables, "EMPLOYEE.RETIRE_DATE", "Retirement Date", aDate, 8, , False, "TERM:RETI-DATE")
'Call AddFieldInfo(VDTables, "EMPLOYEE.ACCRUE_HRS_DAY", "Accrual Hours per Day", aNumber, 12, 2, False, "JH_DHRS")
Call AddFieldInfo(VDTables, "EMPLOYEE.MULT_UNION_FLAG", "Multiple Union Flag", aString, 1, , False, "DFLT:MULT_UNION_FLAG:N")
'?Call AddFieldInfo(VDTables, "EMPLOYEE.HRS_BY_OCC_FLAG", "Track Hours by Occupation", aString, 1, , False)
Call AddFieldInfo(VDTables, "EMPLOYEE.NEXT_MERIT_DATE", "Next Merit Increment Date", aDate, 8, , False, "VIT:Next Merit Date")
Call AddFieldInfo(VDTables, "EMPLOYEE.IN_LIEU_PCT", "In Lieu of Benefit Percent", aNumber, 5, 4, False, "RESET:BF_BCODE")   'Stat Pay % Accrued
Call AddFieldInfo(VDTables, "EMPLOYEE.REHIRE_DATE", "Rehire Date", aDate, 8, , False, "VIT:Last Hire Date")
Call AddFieldInfo(VDTables, "EMPLOYEE.LAID_OFF_DATE", "Laid Off Date", aDate, 8, , False, "TERM:LO-DATE")
'Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_SSN", "Employee Social Security Number", aString, 9, , False,"ED_SSN")
Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_TYPE_CODE", "Employee Type Code", aString, 2, , False, "VIT:Employee Type Code")  'Status
'Call AddFieldInfo(VDTables, "EMPLOYEE.EMP_RPP_HRS", "Employee RPP Hours", aNumber, 12, 2, False)   'Pensionable Hours
'Call AddFieldInfo(VDTables, "EMPLOYEE.FED_JOB_CLASS_CODE", "Federal Job Class Code", aString, 4, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.FED_WT_APPLIC_FLAG", "Federal Withholding Tax Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.FED_WT_ALLOW", "Federal WT Allowances", aNumber, 2, 0, False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.FED_WT_ADD_AMT", "Federal WT Additional Amt", aNumber, 12, 2, False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.STATE_WT_APPLIC_FLAG", "State Tax Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.STATE_WT_ALLOW", "State WT Allowances", aNumber, 2, 0, False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.STATE_WT_ALTER", "State WT Alternate", aNumber, 5, 0, False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.STATE_WT_ADD_AMT", "State WT Additional Amount", aNumber, 12, 2, False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.ELEC_W2_FLAG", "Electronic W2 Flag", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.SS_APPLIC_FLAG", "Social Security Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.MC_APPLIC_FLAG", "Medicare Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.FUTA_APPLIC_FLAG", "FUTA Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.SUI_APPLIC_FLAG", "State EI Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.STATE_WCB_APPLIC_FLAG", "State WCB Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.STATE_PEN_APPLIC_FLAG", "State Pension Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.LOCAL_WT_APPLIC_FLAG", "Local Tax Applicable", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.ACCRUAL_FTE_PCT", "Accrual FTE Percent", aNumber, 5, 4, False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.HOL_PAY_HOUR", "Holiday Pay Hours", aNumber, 12, 4, False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.HOL_PAY_FREQ", "Holiday Pay Frequency", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.FLAT_DIST_CODE", "", aString, 1, , False)
'Call AddFieldInfo(VDTables, "EMPLOYEE.FLAT_DIST_AMT", "", aNumber, 12, 2, False)

'Note:Cannot add CLIENT_CODE from info:HR because Vadim is adding themselves in their stored procedure
'to add new CLIENT record.
'Call AddFieldInfo(VDTables, "CLIENT.CLIENT_CODE", "", aString, 10, , True, "VAD:CLIENT_CODE")
Call AddFieldInfo(VDTables, "CLIENT.CLIENT_TYPE_CODE", "", aString, 1, , True, "CLIENT_DFLT:CLIENT_TYPE_CODE:I")
Call AddFieldInfo(VDTables, "CLIENT.CLIENT_NAME1", "", aString, 25, , False, "ED_SURNAME")
Call AddFieldInfo(VDTables, "CLIENT.CLIENT_NAME2", "", aString, 25, , False, "ED_FNAME")
'Call AddFieldInfo(VDTables, "CLIENT.CLIENT_NAME3", "", aString, 25, , False)
'Call AddFieldInfo(VDTables, "CLIENT.CLIENT_NAME4", "", aString, 25, , False)
Call AddFieldInfo(VDTables, "CLIENT.CLIENT_TITLE1", "", aString, 10, , False, "ED_TITLE")
'Call AddFieldInfo(VDTables, "CLIENT.CLIENT_TITLE2", "", aString, 10, , False)
'Call AddFieldInfo(VDTables, "CLIENT.STREET_NUM", "", aString, 10, , False)
'If glbCompSerial = "S/N - 2378W" Then   'Town of Aurora only
'    Call AddFieldInfo(VDTables, "CLIENT.UNIT_NUM", "", aString, 10, , False, "ED_ADDR2")
'End If
'Call AddFieldInfo(VDTables, "CLIENT.STREET_DIR", "", aString, 2, , False)
Call AddFieldInfo(VDTables, "CLIENT.STREET1", "", aString, 40, , True, "ED_ADDR1")
Call AddFieldInfo(VDTables, "CLIENT.STREET2", "", aString, 40, , False, "ED_ADDR2")
Call AddFieldInfo(VDTables, "CLIENT.CITY", "", aString, 30, , False, "ED_CITY")
Call AddFieldInfo(VDTables, "CLIENT.PROV_STATE_CODE", "", aString, 2, , False, "ED_PROV")
Call AddFieldInfo(VDTables, "CLIENT.COUNTRY_CODE", "", aString, 3, , True, "ED_COUNTRY")
Call AddFieldInfo(VDTables, "CLIENT.POSTAL_ZIP", "", aString, 10, , False, "ED_PCODE")
Call AddFieldInfo(VDTables, "CLIENT.PHONE1", "", aString, 16, , False, "ED_PHONE")
Call AddFieldInfo(VDTables, "CLIENT.PHONE2", "", aString, 16, , False, "ED_BUSNBR")
'Call AddFieldInfo(VDTables, "CLIENT.FAX", "", aString, 16, , False)
Call AddFieldInfo(VDTables, "CLIENT.CLIENT_NET_ADDR", "", aString, 40, , False, "ED_EMAIL")
Call AddFieldInfo(VDTables, "CLIENT.CONTACT_NAME", "", aString, 40, , False, "ED_ECONT")
Call AddFieldInfo(VDTables, "CLIENT.CONTACT_PHONE1", "", aString, 16, , False, "ED_ENBR")
Call AddFieldInfo(VDTables, "CLIENT.CONTACT_PHONE2", "", aString, 16, , False, "ED_EP2NBR")
'?Call AddFieldInfo(VDTables, "CLIENT.CONTACT_FAX", "", aString, 16, , False)
Call AddFieldInfo(VDTables, "CLIENT.CONTACT_NET_ADDR", "", aString, 40, , False, "ED_EEMAIL")
'Call AddFieldInfo(VDTables, "CLIENT.GST_REG_NUM", "", aString, 20, , False)
Call AddFieldInfo(VDTables, "CLIENT.CLIENT_DESC", "", aString, 255, , False, "VAD:CLIENT_DESC")
Call AddFieldInfo(VDTables, "CLIENT.TERM_DATE", "", aDate, 8, , False, "TERM:CLIENT_DATE")
'Call AddFieldInfo(VDTables, "CLIENT.NEW_CLIENT_CODE", "", aString, 10, , False)
'Call AddFieldInfo(VDTables, "CLIENT.INS_USER", "", aString, 30, , True)
'Call AddFieldInfo(VDTables, "CLIENT.INS_DATE", "", aString, 8, , True)
'Call AddFieldInfo(VDTables, "CLIENT.UPD_USER", "", aString, 30, , True)
'Call AddFieldInfo(VDTables, "CLIENT.UPD_DATE", "", aString, 8, , True)
'''
Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.COMPANY_CODE", "", aString, 2, , True, "BANK:COMP_CODE")
Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.EMP_NUM", "", aString, 10, , True, "BANK:PAYROLL_ID")
Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.BANK_CODE", "", aString, 10, , True, "BANK:BANK")
Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.BANK_BRANCH_CODE", "", aString, 6, , True, "BANK:BRANCH")
Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.EMP_BANK_ACT", "", aString, 16, , True, "BANK:ACCOUNT")
Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.DIST_TYPE_CODE", "", aString, 1, , True, "BANK:DISTTYPE")
Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.DISTRIBUTION_VALUE", "", aNumber, 12, 4, True, "BANK:DISTVALUE")
'Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.INS_USER", "", aString, 30, , True)
'Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.INS_DATE", "", aString, 8, , True)
'Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.UPD_USER", "", aString, 30, , True)
'Call AddFieldInfo(VDTables, "PAY_DEPOSIT_DISTRIBUTION.UPD_DATE", "", aString, 8, , True)

'Hemu - New
'Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.COMPANY_CODE", "", aString, 2, , True, "TRANSCODE:COMP_CODE")
'Hemu
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.EMP_NUM", "", aNumber, 12, 4, True, "TRANSCODE:PAYROLL_ID")
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.PAY_CODE", "", aNumber, 12, 4, True, "TRANSCODE:PAYCODE")
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.PAY_CYCLE_FREQ_CODE", "", aNumber, 12, 4, True, "TRANSCODE:FREQCODE")
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.REMAINING_CALC", "", aNumber, 12, 4, False, "TRANSCODE:REMAININGCALC")
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.RATE_LEVEL", "", aNumber, 12, 4, False, "TRANSCODE:RATELEVEL")
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.TRANS_VALUE", "", aNumber, 12, 4, True, "TRANSCODE:TRANSVALUE")
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.EMPLR_MULTIPLIER", "", aNumber, 12, 4, True, "TRANSCODE:EMPLR_MULTIPLIER")
Call AddFieldInfo(VDTables, "EMP_TRANS_CODE.TRANS_DATE", "", aDate, 12, 4, True, "TRANSCODE:DATE")

'Hemu - New
'Call AddFieldInfo(VDTables, "EMP_TRANS.COMPANY_CODE", "", aString, 2, , True, "TRANS:COMP_CODE")
'Hemu
Call AddFieldInfo(VDTables, "EMP_TRANS.EMP_NUM", "", aNumber, 12, 4, True, "TRANS:PAYROLL_ID")
Call AddFieldInfo(VDTables, "EMP_TRANS.EMP_ACCT_NUM", "", aNumber, 12, 4, True, "TRANS:GLNO")
Call AddFieldInfo(VDTables, "EMP_TRANS.OCC_CODE", "", aNumber, 12, 4, True, "TRANS:JOBCODE")
Call AddFieldInfo(VDTables, "EMP_TRANS.PAYCODE", "", aNumber, 12, 4, True, "TRANS:PAYCODE")
Call AddFieldInfo(VDTables, "EMP_TRANS.TRANS_DATE", "", aDate, 12, 4, True, "TRANS:DATE")
Call AddFieldInfo(VDTables, "EMP_TRANS.TRANS_CREATE_TYPE", "", aNumber, 12, 4, True, "TRANS:CREATE_TYPE")
Call AddFieldInfo(VDTables, "EMP_TRANS.TRANS_HOURS", "", aNumber, 12, 4, True, "TRANS:HOURS")
Call AddFieldInfo(VDTables, "EMP_TRANS.TRANS_RATE", "", aNumber, 12, 4, True, "TRANS:RATE")
Call AddFieldInfo(VDTables, "EMP_TRANS.TRANS_AMOUNT", "", aNumber, 12, 4, True, "TRANS:AMOUNT")
Call AddFieldInfo(VDTables, "EMP_TRANS.TRANS_DESC", "", aNumber, 12, 4, True, "TRANS:DESC")

'Hemu - New
'Call AddFieldInfo(VDTables, "OCCUPATION.COMPANY_CODE", "", aString, 2, , True, "OCCUPATION:COMP_CODE")
'Hemu
Call AddFieldInfo(VDTables, "OCCUPATION.OCC_CODE", "", aString, 6, , True, "OCCUPATION:OCC_CODE")
Call AddFieldInfo(VDTables, "OCCUPATION.OCC_NAME", "", aString, 40, , True, "OCCUPATION:OCC_NAME")
Call AddFieldInfo(VDTables, "OCCUPATION.OCC_DESC", "", aString, 255, , False, "OCCUPATION:OCC_DESC")
Call AddFieldInfo(VDTables, "OCCUPATION.OCC_LEVEL_RATE", "", aNumber, 12, 4, True, "OCCUPATION:OCC_LEVEL_RATE")
Call AddFieldInfo(VDTables, "OCCUPATION.RATE_BASED_ON", "", aString, 40, , True, "OCCUPATION:RATE_BASED_ON")
Call AddFieldInfo(VDTables, "OCCUPATION.SALARY_AMOUNT", "", aNumber, 12, 4, True, "OCCUPATION:SALARY_AMOUNT")

'Hemu - New
'Call AddFieldInfo(VDTables, "OCC_RATE.COMPANY_CODE", "", aString, 2, , True, "OCCRATE:COMP_CODE")
'Hemu
Call AddFieldInfo(VDTables, "OCC_RATE.OCC_CODE", "", aString, 12, 4, True, "OCCRATE:OCC_CODE")
Call AddFieldInfo(VDTables, "OCC_RATE.OCC_LEVEL", "", aNumber, 12, 4, True, "OCCRATE:OCC_LEVEL")
Call AddFieldInfo(VDTables, "OCC_RATE.OCC_LEVEL_NAME", "", aString, 40, , True, "OCCRATE:OCC_LEVEL_NAME")
Call AddFieldInfo(VDTables, "OCC_RATE.OCC_RATE", "", aNumber, 12, 4, True, "OCCRATE:OCC_RATE")
Call AddFieldInfo(VDTables, "OCC_RATE.OCC_START_HOUR", "", aNumber, 12, 4, True, "OCCRATE:OCC_START_HOUR")
Call AddFieldInfo(VDTables, "OCC_RATE.OCC_SALARY", "", aNumber, 12, 4, True, "OCCRATE:OCC_SALARY")

'Hemu - New for Campbell River to pass VCH Pin#
Call AddFieldInfo(VDTables, "VCH_ID.APPL_CODE", "", aString, 2, , True, "VCHID:APPL_CODE")
Call AddFieldInfo(VDTables, "VCH_ID.ACCT_SECT_1", "", aString, 3, , True, "VCHID:ACCT_SECT_1")
Call AddFieldInfo(VDTables, "VCH_ID.ACCT_SECT_2", "", aString, 25, , True, "VCHID:ACCT_SECT_2")
Call AddFieldInfo(VDTables, "VCH_ID.ACCT_SECT_3", "", aString, 3, , True, "VCHID:ACCT_SECT_3")
Call AddFieldInfo(VDTables, "VCH_ID.VCH_PIN", "VCH Pin #", aString, 20, , True, "VCHID:VCH_PIN")


End Sub

Sub AddFieldInfo(AddCollection As Collection, vField, vDesc, Optional vType, Optional vLength, Optional vDecimal, Optional Required, Optional xIHR)
    Dim xField As New FieldInfo
    Dim xTemp
    'Dim vTable, vField
    'xTemp = Split(vTableField, ".")
    'vTable = xTemp(0)
    'vField = xTemp(1)
    
    'xField.fdTable = vTable
    xField.fdName = vField
    xField.fdDesc = vDesc
    xField.fdType = vType
    xField.fdLength = vLength
    If Not IsMissing(vDecimal) Then xField.fdDecimal = vDecimal
    xField.fdREQ = Required
    xField.fdIHR = xIHR
    AddCollection.Add xField, vField
End Sub

Sub AddMap()
Dim X, Y, I
Dim HRField
Dim xMap() As String
Dim vdTotal
Dim VDField
Dim FindExist As Boolean
vdTotal = VDTables.count

ReDim xMap(vdTotal, 2)
I = 1
For X = 1 To vdTotal
    HRField = VDTables(X).fdIHR
    VDField = VDTables(X).fdName
    FindExist = False
    For Y = 1 To vdTotal
        If xMap(Y, 1) = HRField Then
            xMap(Y, 2) = xMap(Y, 2) & "," & VDField
            FindExist = True
            Exit For
        End If
    Next
    If Not FindExist Then
        xMap(I, 1) = HRField
        xMap(I, 2) = VDField
        I = I + 1
    End If
Next
For X = 1 To IVMap.count
    IVMap.Remove 1
Next

For X = 1 To vdTotal
    If xMap(X, 1) <> "" Then
        xMap(X, 2) = xMap(X, 1) & "," & xMap(X, 2)
        IVMap.Add xMap(X, 2), xMap(X, 1)
    End If
Next
Vadim_PayType_field = getVITField("Payment Type")
Vadim_EmpType_field = getVITField("Employee Type Code")
Vadim_PayType_TABLName = getTABLName(Vadim_PayType_field)
Vadim_EmpType_TABLName = getTABLName(Vadim_EmpType_field)
End Sub


Function AddBatchVadim(UptType, Optional UptDate)
Dim rsBatch As New ADODB.Recordset
Dim ADDBatchID
If gdbPayroll Is Nothing Then Exit Function
If gdbPayroll.ConnectionString = "" Then Exit Function

If glbVadim Then
    rsBatch.Open "SELECT * FROM SY_INTERFACE_BATCH", gdbPayroll, adOpenKeyset, adLockOptimistic
    rsBatch.AddNew
    rsBatch("SOURCE_APP") = "HR"
    rsBatch("PROCESS_TYPE") = IIf(UptType = "T", "M", UptType)
    rsBatch("PROCESS_CODE") = "I"
    rsBatch("PROCESS_USER") = glbUserID
    If IsDate(UptDate) Then
        rsBatch("PROCESS_DATE") = UptDate
    Else
        rsBatch("PROCESS_DATE") = Date
    End If
    rsBatch.Update
    ADDBatchID = rsBatch("SY_BATCH_ID").Value
    AddBatchVadim = ADDBatchID
End If
End Function

Sub TermVadimEmp(UptDate, xEmpnbr, xPayID, Optional PStatus As PassStatus)
Dim HRFields As New Collection
Dim xBatchID
Dim X, Y
Dim PayIDs
Dim xPayment_oldValue
Dim xPayment_newValue
Dim rsHREmp As New ADODB.Recordset
Dim xOEmail As String

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(Termination) Then Exit Sub
If Not IsDate(glbChgTermDate) Then Exit Sub
xBatchID = AddBatchVadim("M")

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    'For City of Kawartha Lakes - If employee does not get any benefits then they get
    'terminated with Payment Type "T" and Termination date. An employee is marked LayOff or
    'Retire when they get benefits and so not terminated.
    HRFields.Add "TERM:EMP_DATE"
    xPayment_newValue = "T"
Else
    If (glbChgTermReason = "LO" Or glbChgTermReason = "LAYO") Then
        HRFields.Add "TERM:LO-DATE"
        HRFields.Add "TERM:EMP_DATE"
        xPayment_newValue = "L"
    ElseIf glbChgTermReason = "RETI" Then
        HRFields.Add "TERM:RETI-DATE"
        HRFields.Add "TERM:EMP_DATE"
        
        'Ticket #25363 - County of Lambton wants to transfer T for RETIRE
        If glbCompSerial = "S/N - 2355W" Then
            xPayment_newValue = "T"
        Else
            xPayment_newValue = "R"
        End If
    Else
        HRFields.Add "TERM:EMP_DATE"
        xPayment_newValue = "T"
    End If
End If

If Vadim_PayType_field <> "" Then
    xPayment_oldValue = getEmpValue(Vadim_PayType_field, xEmpnbr, xPayID)
End If

PayIDs = Split(getPayrollIDs(xEmpnbr, xPayID), "|")
For X = 0 To UBound(PayIDs)
    xBatchID = AddBatchVadim("M", UptDate)
    For Y = 1 To HRFields.count
        Call VadimInterface(xBatchID, PayIDs(X), HRFields(Y), Null, glbChgTermDate)
    Next
    If Vadim_PayType_field <> "" Then
        Call VadimInterface(xBatchID, PayIDs(X), Vadim_PayType_field, xPayment_oldValue, xPayment_newValue)
    End If
    'Ticket #24996 - City of Campbell River
    'If glbCompSerial = "S/N - 2458W" Then
    '    Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_ACTIVE_FLAG:Y", "1", "2")
    'Else
        Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_ACTIVE_FLAG:Y", "Y", "N")
    'End If
    Call CloseBatchVadim(xBatchID)
Next

'Ticket #29480 - Update Employee History for this Payment Type change
If Vadim_PayType_field = "ED_LOC" Then
    If Not EmpHisCalc(2, xEmpnbr, "", "", "", "", "", "", "", Date, "LOC", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
ElseIf Vadim_PayType_field = "ED_SECTION" Then
    If Not EmpHisCalc(2, xEmpnbr, "", "", "", "", "", "", "", Date, "SECTION", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
ElseIf Vadim_PayType_field = "ED_ADMINBY" Then
    If Not EmpHisCalc(2, xEmpnbr, "", "", "", "", "", "", "", Date, "ADMINBY", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
ElseIf Vadim_PayType_field = "ED_REGION" Then
    If Not EmpHisCalc(2, xEmpnbr, "", "", "", "", "", "", "", Date, "REGION", xPayment_newValue, , , , , xPayment_oldValue) Then MsgBox "EMPHIS Error "
End If

'County of Lambton - Ticket #14971, Only from Termination screen
If glbCompSerial = "S/N - 2355W" And PStatus = Termination Then
    'Update Employee's ED_EMAIL field to blank
    xOEmail = ""
    rsHREmp.Open "SELECT ED_EMPNBR, ED_EMAIL FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHREmp.EOF Then
        xOEmail = IIf(IsNull(rsHREmp("ED_EMAIL")), "", rsHREmp("ED_EMAIL"))
        rsHREmp("ED_EMAIL") = ""
        rsHREmp.Update
        
        For X = 0 To UBound(PayIDs)
            xBatchID = AddBatchVadim("M", UptDate)
            Call VadimInterface(xBatchID, PayIDs(X), "ED_EMAIL", xOEmail, "")
            Call CloseBatchVadim(xBatchID)
        Next
    End If
    rsHREmp.Close
    Set rsHREmp = Nothing
End If

End Sub


Sub ReHireVadimEmp(UptDate, xEmpnbr, xPayID)
Dim HRFields As New Collection
Dim xBatchID
Dim X, Y
Dim PayIDs
Dim rsTerm As New ADODB.Recordset
Dim rsPA As New ADODB.Recordset
Dim xPayment_newValue
Dim xPayment_oldValue

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(Rehire) Then Exit Sub
rsTerm.Open "SELECT * FROM TERM_HRTRMEMP WHERE Employee_Number=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly
If Not rsTerm.EOF Then
    glbChgTermReason = rsTerm("TERM_REASON")
    glbChgTermDate = rsTerm("TERM_DOT")
End If
rsTerm.Close
xBatchID = AddBatchVadim("M")

'City of Kawartha Lakes
If glbCompSerial = "S/N - 2363W" Then
    'For City of Kawartha Lakes - If employee does not get any benefits then they get
    'terminated with Payment Type "T" and Termination date. An employee is marked LayOff or
    'Retire when they get benefits and so not terminated.
    HRFields.Add "TERM:EMP_DATE"
    xPayment_newValue = "T"
Else
    If (glbChgTermReason = "LO" Or glbChgTermReason = "LAYO") Then
        HRFields.Add "TERM:LO-DATE"
        HRFields.Add "TERM:EMP_DATE"
        xPayment_oldValue = "L"
    ElseIf glbChgTermReason = "RETI" Then
        HRFields.Add "TERM:RETI-DATE"
        HRFields.Add "TERM:EMP_DATE"
        xPayment_oldValue = "R"
    Else
        HRFields.Add "TERM:EMP_DATE"
        xPayment_oldValue = "T"
    End If
End If

If Vadim_PayType_field <> "" Then
    xPayment_newValue = getEmpValue(Vadim_PayType_field, xEmpnbr, xPayID)
End If

PayIDs = Split(getPayrollIDs(xEmpnbr, xPayID, True), "|")
For X = 0 To UBound(PayIDs)
    xBatchID = AddBatchVadim("M", UptDate)
    For Y = 1 To HRFields.count
        'City of Timmins - Vadim is not accepting "Null" or Null for Term date to be cleared
        If glbCompSerial = "S/N - 2375W" And HRFields(Y) = "TERM:EMP_DATE" Then
            Call VadimInterface(xBatchID, PayIDs(X), HRFields(Y), glbChgTermDate, "")
        Else
            Call VadimInterface(xBatchID, PayIDs(X), HRFields(Y), glbChgTermDate, Null)
        End If
    Next
    If Vadim_PayType_field <> "" Then
        Call VadimInterface(xBatchID, PayIDs(X), Vadim_PayType_field, xPayment_oldValue, xPayment_newValue)
    End If
    'Ticket #24996 - City of Campbell River
    'If glbCompSerial = "S/N - 2458W" Then
    '    Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_ACTIVE_FLAG:Y", "2", "1")
    'Else
        Call VadimInterface(xBatchID, PayIDs(X), "DFLT:EMP_ACTIVE_FLAG:Y", "N", "Y")
    'End If
    
    'Ticket #25412 - Town of Greater Napanee
    If glbCompSerial = "S/N - 2447W" Then
        'Do not update Seniority Date in Vadim
        'Update Start Date and EI Start Date in Vadim with Original Hire Date
        'Call VadimInterface(xBatchID, PayIDs(X), "VIT:Start Date", "", GetEmpData_PayrollID(PayIDs(X), "ED_DOH"))  'Already being passed
        Call VadimInterface(xBatchID, PayIDs(X), "VIT:EI Start Date", "", GetEmpData_PayrollID(PayIDs(X), "ED_DOH"))
    Else
        'Ticket #29007 - City of Campbell River don't want Seniority to transfer to Vadim
        If glbCompSerial <> "S/N - 2458W" Then
            Call VadimInterface(xBatchID, PayIDs(X), "ED_SENDTE", "", GetEmpData_PayrollID(PayIDs(X), "ED_SENDTE"))
        End If
    End If
    
    'City of Timmins - Pass TD1 and Provincial Amounts as well
    If glbCompSerial = "S/N - 2375W" Then
        rsPA.Open "select PC_NEXT_AVAILABLE_NBR,PC_FEDTAX,PC_PROVTAX from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
        If Not rsPA.EOF Then
            Call VadimInterface(xBatchID, PayIDs(X), "ED_TD1DOL", "", rsPA("PC_FEDTAX"))
            Call Passing_PAYINFO_Vadim(PayIDs(X), "ED_PROVAMT", "", rsPA("PC_PROVTAX"))
        End If
        rsPA.Close
    End If
    
    Call CloseBatchVadim(xBatchID)
Next
End Sub

Sub DeleteVadimEmp(UptDate, xEmpnbr, Optional xPayID)
Dim HRField
Dim xBatchID
Dim X, PayIDs
HRField = "JH_PAYROLL_ID"
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(Demographices) Then Exit Sub
PayIDs = Split(getPayrollIDs(xEmpnbr, xPayID), "|")
For X = 0 To UBound(PayIDs)
    xBatchID = AddBatchVadim("D", UptDate)
    Call VadimInterface(xBatchID, PayIDs(X), HRField, Null, Null)
    Call CloseBatchVadim(xBatchID)
Next
End Sub

Sub AddNewVadimEmp(PStatus As PassStatus, UptDate, xEmpnbr, xPayID)
Dim rsEmp As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim rsBank As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim rsBNT As New ADODB.Recordset
Dim X
Dim xValue
Dim xMapField
Dim xFieldList
Dim xField As ADODB.Field
Dim IHRFields
Dim Banks As New Collection
Dim xBatchID
Dim HRChanges As New Collection
Dim VCHID As New Collection
Dim NumComp As Boolean
Dim xHRChange As New HRChange
Dim xClientCode As String
Dim SQLQ As String

On Error GoTo AddNewVadimEmp_err

If IsMissing(xEmpnbr) Then xEmpnbr = glbLEE_ID

If Not isTransfer(PStatus) Then Exit Sub

If PStatus = Rehire Then
    Call ReHireVadimEmp(UptDate, xEmpnbr, xPayID)
    Exit Sub
    'do not need to add the new employee, just remove the terminaiton date
End If

'----------------adding records to CLIENT table
' vadim has asked to transfer the client table prior the employee table
Set HRChanges = New Collection
rsEmp.Open "SELECT " & getFieldList(CLIENT, "HREMP") & " FROM HREMP WHERE ED_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly

''Dist. of Muskoka wants to create their own Client Code instead of vadim creating it.
'But cannot be done as Vadim is adding themselves in the stored procedure to add new client record.
'Hence giving duplicate field error in the SET command.
'If glbCompSerial = "S/N - 2373W" Then
'    xClientCode = "Z" & xEmpnbr & "001"
'    Call Add_DFLT_Values(HRChanges, "VAD:CLIENT_CODE", xClientCode)
'End If

For Each xField In rsEmp.Fields
    NumComp = xField.Type = adNumeric Or xField.Type = adInteger
    Call isChanged_Field(HRChanges, Null, xField, NumComp)
Next
rsEmp.Close

'adding default
xHRChange.HRField = "CLIENT_DFLT:CLIENT_TYPE_CODE:I"
xHRChange.NewValue = "I"
xHRChange.OldValue = Null
HRChanges.Add xHRChange, "CLIENT_DFLT:CLIENT_TYPE_CODE:I"

'Ticket #25469 - City of Campbell River - Transfer Client Description
If glbCompSerial = "S/N - 2458W" Then
    Call Add_DFLT_Values(HRChanges, "VAD:CLIENT_DESC", "PA")
End If

Call Passing_Changes_Vadim(HRChanges, PStatus, "A", UptDate, xEmpnbr, xPayID)

'----------------adding records to EMPLOYEE table
Set HRChanges = New Collection
'Ticket #29498 - Commented this because added the following line under SQLQ string
'rsEmp.Open "SELECT " & getFieldList(EMPLOYEE, "HREMP") & " FROM HREMP WHERE ED_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly
SQLQ = "SELECT " & getFieldList(EMPLOYEE, "HREMP")
'Ticket #29498 - City of Campbell River - The Employee Title (ED_ORGT1) also needs to be transferred on New Hire
If glbCompSerial = "S/N - 2458W" Then
    SQLQ = SQLQ & ",ED_ORGT1 "
End If
SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR=" & xEmpnbr
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
For Each xField In rsEmp.Fields
    NumComp = xField.Type = adNumeric Or xField.Type = adInteger
    If xField.name = "ED_PAYROLL_ID" And xPayID <> "" Then
        Call Add_DFLT_Values(HRChanges, "JH_PAYROLL_ID", xPayID)
    Else
        Call isChanged_Field(HRChanges, Null, xField, NumComp)
    End If
Next
rsEmp.Close

'adding default
For X = 1 To IVMap.count
    IHRFields = Split(IVMap(X), ",")
    xMapField = Split(IHRFields(0), ":")
    If xMapField(0) = "DFLT" Then
        If (glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2276W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W" Or glbCompSerial = "S/N - 2373W") And xMapField(1) = "PROBATION_DATE" Then
            'Town of Aurora or City of Kawartha Lakes or City of Niagara Falls or Town of Lasalle, Town of Greater Napanee(Ticket #24375), City of Campbell River(Ticket #24996)-they do not user this, DMuskoka
            'Do not pass the default value at this moment. it wil be added with ED_DOH
        ElseIf (glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2355W" Or glbCompSerial = "S/N - 2276W" Or glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W") And xMapField(1) = "EMP_CLASS_CODE" Then
            'Do not pass the Employee Class Value for City of Kawartha Lakes, City of Timmins, County of Lambton, Ciy of Niagara Falls, Town of Aurora, Dist. of Muskoka, Town of Lasalle, Town of Greater Napanee(Ticket #24375), City of Campbell River(Ticket #24996)
        'ElseIf glbCompSerial = "S/N - 2458W" And xMapField(1) = "VOL_FIREMAN_FLAG" Then    'Mandatory for New Hire cannot be Null
            'Do not pass
        Else
            If glbCompSerial = "S/N - 2363W" And xMapField(1) = "VOL_FIREMAN_FLAG" Then   'City of Kawartha Lakes
                xValue = GetEmpData_PayrollID(xPayID, "ED_LOC")
                If xValue <> "F" Then
                    xValue = xMapField(2)
                    Call Add_DFLT_Values(HRChanges, IHRFields(0), xValue)
                End If
            'ElseIf glbCompSerial = "S/N - 2458W" And xMapField(1) = "EMP_ACTIVE_FLAG" Then
            '    'Ticket #24996 - City of Campbell River
            '    xValue = xMapField(2)
            '    If xValue = "Y" Then xValue = "1" Else xValue = "2"
            '    Call Add_DFLT_Values(HRChanges, IHRFields(0), xValue)
            Else
                xValue = xMapField(2)
                Call Add_DFLT_Values(HRChanges, IHRFields(0), xValue)
            End If
        End If
    End If
Next

If PStatus = Demographices Then
    Call Add_DFLT_Values(HRChanges, "ED_DDI", "D")
    Call Add_DFLT_Values(HRChanges, "ED_DOH", Date)
End If

Call Add_DFLT_Values(HRChanges, "JH_PAYROLL_CATEGORY", getDeftPayCat)

'Ticket #23795 - Town of Lasalle - Pay Frequency - Default to 52
If glbCompSerial = "S/N - 2379W" Then
    Call Add_DFLT_Values(HRChanges, "SH_PAYP", "52")
ElseIf glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W" Or glbCompSerial = "S/N - 2373W" Then
    'Ticket #24375 - Town of Greater Napanee - Default to 26
    'Ticket #24996 - City of Campbell River - Default to 26
    'Ticket #24565 - City of Niagara Falls - Default to 26
    Call Add_DFLT_Values(HRChanges, "SH_PAYP", "26")
Else
    Call Add_DFLT_Values(HRChanges, "SH_PAYP", getDeftPayPeriod)
End If

'Ticket #24018 - Town of Lasalle - Does not like blank value in Vadim Stored Procedure - gets stuck in a loop
'Town of Lasalle - required this as SUPERAN_APPLIC_FLAG cannot be null - this will assign N and then from Banking
'screen it will assign the right value due to Pension Code
'Ticket #23795 - Town of Lasalle - They are not transferring OMERS date. OMERS is dependent on Pension Code.
If glbCompSerial = "S/N - 2379W" Then
    Call Add_DFLT_Values(HRChanges, "ED_OMERS", "N")
Else
    Call Add_DFLT_Values(HRChanges, "ED_OMERS", "")
End If
Call Add_DFLT_Values(HRChanges, "ED_TD1DOL", "0")
Call Add_DFLT_Values(HRChanges, "ED_TD3", "0")
Call Add_DFLT_Values(HRChanges, "ED_VACPC", "0")
Call Add_DFLT_Values(HRChanges, "JH_DHRS", "0")
Call Add_DFLT_Values(HRChanges, "ED_EMPTYPE", "N")

'Ticket #23795 - Town of Lasalle - More default values
If glbCompSerial = "S/N - 2379W" Then
    'Call Add_DFLT_Values(HRChanges, "ED_VADIM1", "0")   'Overhead %    'gives multiple set values
    'Call Add_DFLT_Values(HRChanges, "ED_VADIM2", "0")   'Stat Pay % Paid   'gives multiple set values
    Call Add_DFLT_Values(HRChanges, "RESET:BF_BCODE", "0")  'Stat Pay % Accrued
End If

'Town of Greater Napanee - Ticket #24375
'Ticket #24996 - City of Campbell River
If glbCompSerial = "S/N - 2447W" Or glbCompSerial = "S/N - 2458W" Then
    Call Add_DFLT_Values(HRChanges, "RESET:BF_BCODE", "0")  'Stat Pay % Accrued
End If

Call Passing_Changes_Vadim(HRChanges, PStatus, "A", UptDate, xEmpnbr, xPayID)

'City of Timmins or City of Kawartha Lakes or Town of Aurora or Dist. of Muskoka, Town of Lasalle, Town of Greater Napanee - Ticket #24375
If glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Then
    Call Add_EHT(xPayID)
End If

'----------------adding records to PAY_DEPOSIT_DISTRIBUTION table
Set Banks = New Collection
rsBank.Open "SELECT " & getFieldList(PAY_DEPOSIT_DISTRIBUTION) & "  FROM HREMP WHERE ED_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly
For X = 0 To rsBank.Fields.count - 1
    Call Add_Bank_Info(Banks, "", rsBank.Fields(X))
Next
rsBank.Close
Call Passing_Bank_Vadim(Banks, xEmpnbr, xPayID)

'---------------adding VCH Pin # to VCH_ID table
'Ticket #25469 - City of Campbell River - Transfer VCH Pin #
If glbCompSerial = "S/N - 2458W" Then
    Call Add_VCH_ID(xPayID)
End If

'''rehire doesn't need to add new employee in vadim
'''If PStatus = Rehire Then
'''    If Not isTransfer(Rehire) Then Exit Sub
'''    '----------------MODIFYING records to EMPLOYEE table FROM HR_JOB_HISTORY
'''    rsJOB.Open "SELECT " & getFieldList(EMPLOYEE, "HR_JOB_HISTORY") & " FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly
'''    Do Until rsJOB.EOF
'''        xPayID = Format(rsJOB("JH_PAYROLL_ID"), "@")
'''        If xPayID <> "" Then ' DO NOT NEED PASS TO VADIM BECAUSE THERE IS NO PAYID SETUP
'''            xBatchID = AddBatchVadim("M")
'''            For Each xField In rsJOB.Fields
'''                xValue = xField.Value
'''                If glbLambton Then
'''                    If xField.name = "JH_JOB" And Format(rsJOB("JH_GRID"), "@") <> "" Then
'''                        xValue = Left(rsJOB("JH_GRID"), 1) & xValue & Mid(rsJOB("JH_GRID"), 2)
'''                    End If
'''                End If
'''                If Not (IsNull(xValue) Or IsEmpty(xValue)) Then
'''                    Call VadimInterface(xBatchID, xPayID, xField.name, Null, xValue)
'''                End If
'''            Next
'''            Call CloseBatchVadim(xBatchID)
'''        End If
'''        rsJOB.MoveNext
'''    Loop
'''    rsJOB.Close
'''    '----------------MODIFYING records to EMPLOYEE table FROM HR_JOB_HISTORY
'''    rsSAL.Open "SELECT " & getFieldList(EMPLOYEE, "HR_JOB_HISTORY") & ",SH_SALARY,SH_SALCD  FROM HR_SALARY_HISTORY WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly
'''    Do Until rsSAL.EOF
'''        xPayID = Format(rsSAL("SH_PAYROLL_ID"), "@")
'''        If xPayID <> "" Then ' DO NOT NEED PASS TO VADIM BECAUSE THERE IS NO PAYID SETUP
'''            xBatchID = AddBatchVadim("M")
'''            For Each xField In rsSAL.Fields
'''                xValue = xField.Value
'''                If Not (IsNull(xValue) Or IsEmpty(xValue)) Then
'''                    Call VadimInterface(xBatchID, xPayID, xField, Null, xValue)
'''                End If
'''            Next
'''            Call CloseBatchVadim(xBatchID)
'''            'WILL ADD EMP_TRANS_CODE TO HERE FOR SALARY CHANGE
'''        End If
'''        rsSAL.MoveNext
'''    Loop
'''    rsSAL.Close
'''End If

Exit Sub

AddNewVadimEmp_err:
glbFrmCaption$ = "Add new employee to Vadim"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Add new Emp", "Vadim", "Update")
Call RollBack '21June99 js
Resume Next
End Sub

Sub Add_EHT(xPayID)
Dim X
Dim VDClause
Dim UptType
Dim PayCodeInfo As PayCodeInfoType
Dim PAYINFOBatchID
If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

If Not isTransfer(Demographices) Then Exit Sub

UptType = "A"
VDClause = ""

PayCodeInfo.PayCode = "EHT"
PayCodeInfo.PayType = ""
PayCodeInfo.PayTypeID = ""
PayCodeInfo.PayFreq = "P"


If Not isExistTransCode(xPayID, PayCodeInfo.PayCode) = 1 Then
    PAYINFOBatchID = AddBatchVadim("A")
    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYROLL_ID", "", xPayID)
    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:PAYCODE", "", PayCodeInfo.PayCode)
    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:FREQCODE", "", PayCodeInfo.PayFreq)
    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:REMAININGCALC", "", "9999")
    'City of Timmins or City of Kawartha Lakes or Town of Aurora or Dist. of Muskoka, Town of Lasalle, Town of Greater Napanee - Ticket #24375
    If glbCompSerial = "S/N - 2375W" Or glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2378W" Or glbCompSerial = "S/N - 2373W" Or glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2447W" Then
        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", "", "Null")
    Else
        Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:RATELEVEL", "", 0)
    End If
    
    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:TRANSVALUE", "", 0)
    Call VadimInterface(PAYINFOBatchID, xPayID, "TRANSCODE:EMPLR_MULTIPLIER", "", 0)
    Call CloseBatchVadim(PAYINFOBatchID)
End If
End Sub

Sub Add_VCH_ID(xPayID)
Dim X
Dim VDClause
Dim UptType
Dim VCHIDBatchID

If gdbPayroll Is Nothing Then Exit Sub
If gdbPayroll.ConnectionString = "" Then Exit Sub

UptType = "A"
VDClause = ""

If Not isExistVCHID(xPayID) = 1 Then
    VCHIDBatchID = AddBatchVadim("A")
    Call VadimInterface(VCHIDBatchID, xPayID, "VCHID:APPL_CODE", "", "PA")
    Call VadimInterface(VCHIDBatchID, xPayID, "VCHID:ACCT_SECT_1", "", "01")
    Call VadimInterface(VCHIDBatchID, xPayID, "VCHID:ACCT_SECT_2", "", xPayID)
    Call VadimInterface(VCHIDBatchID, xPayID, "VCHID:ACCT_SECT_3", "", "")
    Call VadimInterface(VCHIDBatchID, xPayID, "VCHID:VCH_PIN", "", xPayID)
    Call CloseBatchVadim(VCHIDBatchID)
End If
End Sub


Sub CloseBatchVadim(CloseBatchID, Optional ProcessCode)
Dim rsBatch As New ADODB.Recordset
Dim rsInter As New ADODB.Recordset

rsBatch.Open "SELECT * FROM SY_INTERFACE_BATCH WHERE SY_BATCH_ID=" & CloseBatchID, gdbPayroll, adOpenStatic, adLockOptimistic
rsInter.Open "SELECT SY_BATCH_ID,SY_INTERFACE_ID FROM SY_INTERFACE WHERE SY_BATCH_ID=" & CloseBatchID, gdbPayroll, adOpenForwardOnly
If Not rsBatch.EOF Then
    If rsInter.EOF Then
        rsBatch.Delete
    Else
        On Error GoTo err_update_process_code
        gdbPayroll.BeginTrans
        If Not IsMissing(ProcessCode) Then  'Ticket #24565 - District Municipality of Muskoka - P - 'Pending' Process Code for Future Dated Salary records
            gdbPayroll.Execute "UPDATE SY_INTERFACE_BATCH SET PROCESS_CODE='" & ProcessCode & "' WHERE SY_BATCH_ID=" & CloseBatchID
        Else
            gdbPayroll.Execute "UPDATE SY_INTERFACE_BATCH SET PROCESS_CODE='N' WHERE SY_BATCH_ID=" & CloseBatchID
        End If
        gdbPayroll.CommitTrans
        GoTo TheEnd
err_update_process_code:
        gdbPayroll.RollbackTrans
        Call ERR_Hndlr(Err, "Vadim_Integration", "Update Process_Code From ""I"" to ""N""", "SY_INTERFACE_BATCH", "UPDATE")
        Resume Next
        Screen.MousePointer = DEFAULT
    End If
End If

TheEnd:
rsBatch.Close
rsInter.Close

End Sub


Function getFieldList(vdTable As VadimTableNum, Optional HRTable As String)
Dim fList
fList = ""
Select Case vdTable
Case EMPLOYEE
    Select Case HRTable
    Case "HR_JOB_HISTORY"
        fList = fList & " JH_PAYROLL_ID"
        fList = fList & ",JH_JOB"
        fList = fList & ",JH_GRID"
        fList = fList & ",JH_GLNO"
        fList = fList & ",JH_DEPTNO"
        fList = fList & ",JH_ORG"
        fList = fList & ",JH_DHRS"
    Case "HR_SALARY_HISTORY"
        fList = fList & " SH_PAYP"
        fList = fList & ",SH_GRID"
        fList = fList & ",SH_NEXTDAT"
        fList = fList & ",SH_EDATE"
        fList = fList & ",SH_GRADE"
    Case "HREMP"
        'Demographics
        fList = fList & " ED_PAYROLL_ID"
        fList = fList & ",ED_SIN"
        fList = fList & ",ED_SSN"
        fList = fList & ",ED_SEX"
        fList = fList & ",ED_MSTAT"
        fList = fList & ",ED_DOB"
        fList = fList & ",ED_DRIVERLIC"
        fList = fList & ",ED_GLNO"
        fList = fList & ",ED_DEPTNO"
        'status
        fList = fList & ",ED_ORG"
        fList = fList & ",ED_EMPTYPE"
        fList = fList & ",ED_SALDIST"
        fList = fList & ",ED_PT"
        'banking
        fList = fList & ",ED_DDI"
        fList = fList & ",ED_TD1DOL"
        fList = fList & ",ED_TD3"
        fList = fList & ",ED_TD3PC"
        fList = fList & ",ED_PROVAMT"
        fList = fList & ",ED_CPP"
        fList = fList & ",ED_OMERS"
        fList = fList & ",ED_UIC"
        fList = fList & ",ED_GROSSCD"
        fList = fList & ",ED_VACPC"
        fList = fList & ",ED_WCB"
        fList = fList & ",ED_WCBCODE"
        'fList = fList & ",ED_VADIM1"
        'vadim import table
        fList = fList & ",ED_DIV"
        fList = fList & ",ED_LOC"
        fList = fList & ",ED_ADMINBY"
        fList = fList & ",ED_REGION"
        fList = fList & ",ED_SECTION"
        fList = fList & ",ED_DOH"
        fList = fList & ",ED_SENDTE"
        fList = fList & ",ED_LTHIRE"
        fList = fList & ",ED_UNION"
        fList = fList & ",ED_FDAY"
        fList = fList & ",ED_LDAY"
        fList = fList & ",ED_USRDAT1"
        fList = fList & ",ED_HOMEOPRTNBR" 'for category default
        
    End Select
Case CLIENT
        fList = fList & " ED_SURNAME"
        fList = fList & ",ED_FNAME"
        fList = fList & ",ED_TITLE"
        fList = fList & ",ED_ADDR1"
        fList = fList & ",ED_ADDR2"
        fList = fList & ",ED_CITY"
        fList = fList & ",ED_PROV"
        fList = fList & ",ED_COUNTRY"
        fList = fList & ",ED_PCODE"
        fList = fList & ",ED_PHONE"
        fList = fList & ",ED_BUSNBR"
        fList = fList & ",ED_EMAIL"
        fList = fList & ",ED_ECONT"
        fList = fList & ",ED_ENBR"
        fList = fList & ",ED_EP2NBR"
        fList = fList & ",ED_EEMAIL"
Case PAY_DEPOSIT_DISTRIBUTION
        fList = fList & " ED_BANK,ED_BRANCH,ED_ACCOUNT,ED_AMTDEPOSIT,ED_PCDEPOSIT"
        fList = fList & ",ED_BANK2,ED_BRANCH2,ED_ACCOUNT2,ED_AMTDEPOSIT2,ED_PCDEPOSIT2"
        fList = fList & ",ED_BANK3,ED_BRANCH3,ED_ACCOUNT3,ED_AMTDEPOSIT3,ED_PCDEPOSIT3"
End Select
getFieldList = fList
End Function

Function isTransfer(PStatus As PassStatus)
Dim rsVadimSetup As New ADODB.Recordset
Dim xIhrItem
isTransfer = False
Select Case PStatus
 
Case Demographices, Banking, Contacts, Position, Salary, Status
    xIhrItem = "Employee Basic Information"
Case Termination
    xIhrItem = "Termination"
Case Rehire
    xIhrItem = "Rehire"
Case Attendance
    xIhrItem = "Attendance"
Case Benefit
    xIhrItem = "Benefit"
Case SalaryGirdMaster
    xIhrItem = "Salary Grid Details"
Case PositionMaster
    xIhrItem = "Position Master"
End Select
rsVadimSetup.Open "SELECT FromHR FROM VADIM_SETUP WHERE IHRItem='" & xIhrItem & "'", gdbAdoIhr001, adOpenForwardOnly
If Not rsVadimSetup.EOF Then
    isTransfer = rsVadimSetup("FromHR") <> 0
End If
rsVadimSetup.Close
End Function
Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, X%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
Round2DEC = Round(tmpNUM, glbCompDecHR)

End Function

Private Function getDeftPayCat()
Dim rsPayCat As New ADODB.Recordset
getDeftPayCat = ""
rsPayCat.Open "SELECT TOP 1 PC_CODE FROM HR_PAYROLL_CATEGORY", gdbAdoIhr001, adOpenForwardOnly
If Not rsPayCat.EOF Then
    getDeftPayCat = rsPayCat("PC_CODE") & ""
End If
rsPayCat.Close
End Function

Public Function getPayType(xEmpnbr)
Dim rsPayType As New ADODB.Recordset
Dim xPayTypeField
getPayType = ""
xPayTypeField = getVITField("Payment Type")
rsPayType.Open "SELECT TOP 1 " & xPayTypeField & " FROM HREMP WHERE ED_EMPNBR=" & xEmpnbr, gdbAdoIhr001, adOpenForwardOnly
If Not rsPayType.EOF Then
    getPayType = rsPayType(xPayTypeField) & ""
End If
rsPayType.Close
End Function

Private Function getDeftPayPeriod()
Dim rsPayPeriod As New ADODB.Recordset
getDeftPayPeriod = ""
rsPayPeriod.Open "SELECT TOP 1 TB_KEY FROM HRTABL WHERE TB_NAME='SDPP'", gdbAdoIhr001, adOpenForwardOnly
If Not rsPayPeriod.EOF Then
    getDeftPayPeriod = rsPayPeriod("TB_KEY") & ""
End If
rsPayPeriod.Close
End Function

Public Function isExistTransCode(xPayID, xPayCode)
Dim rsTransCode As New ADODB.Recordset
On Error GoTo TheEnd
isExistTransCode = 2 'DO NOT IF EXISTS
rsTransCode.Open "SELECT TRANS_VALUE FROM EMP_TRANS_CODE WHERE EMP_NUM='" & xPayID & "' AND PAY_CODE='" & xPayCode & "'", gdbPayroll, adOpenForwardOnly
If rsTransCode.BOF Then
    isExistTransCode = 0
Else
    isExistTransCode = 1
End If
TheEnd:
End Function

Public Function isExistVCHID(xPayID)
Dim rsTransCode As New ADODB.Recordset
On Error GoTo TheEnd
isExistVCHID = 2 'DO NOT IF EXISTS
rsTransCode.Open "SELECT ACCT_SECT_2 FROM VCH_ID WHERE APPL_CODE = 'PA' AND ACCT_SECT_2='" & xPayID & "'", gdbPayroll, adOpenForwardOnly
If rsTransCode.BOF Then
    isExistVCHID = 0
Else
    isExistVCHID = 1
End If
TheEnd:
End Function

Private Function convertCountry(xCountry)
Select Case xCountry
Case "U.S.A."
    xCountry = "USA"
Case Else
    xCountry = Left(xCountry, 3)
End Select

End Function
Private Function ifExistVadimOccCode(ByVal xOccCode As String)
Dim X
Dim xBNo
Dim SQLQ
Dim rsVP As New ADODB.Recordset
On Error GoTo default_this
If xOccCode <> "" Then
    SQLQ = "SELECT OCC_CODE FROM OCCUPATION WHERE OCC_CODE ='" & xOccCode & "'"
    rsVP.Open SQLQ, gdbPayroll, adOpenForwardOnly
    If Not rsVP.EOF Then
        ifExistVadimOccCode = True
    Else
        ifExistVadimOccCode = False
    End If
    rsVP.Close
    Exit Function
End If
default_this:
    ifExistVadimOccCode = False
End Function
Private Function getTABLName(xField)
    Select Case xField
    Case "ED_REGION"
        getTABLName = "EDRG"
    Case "ED_LOC"
        getTABLName = "EDLC"
    Case "ED_ADMINBY"
        getTABLName = "EDAB"
    Case "ED_SECTION"
        getTABLName = "EDSE"
    Case "ED_HIRECODE"
        getTABLName = "EDHC"
    End Select
End Function

Public Function CodeMatrix(TblName As String, TblKey As String, DefaultTo As String) As String
Dim rsCodeMatrix As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM CODEMATRIX WHERE CM_NAME='" & TblName & "' AND CM_KEY='" & TblKey & "'"
    rsCodeMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsCodeMatrix.EOF Then
        CodeMatrix = DefaultTo
    Else
        CodeMatrix = rsCodeMatrix("CM_PRKEY")
    End If
    rsCodeMatrix.Close
End Function

Public Function Get_Hourly_Rate_NiagaraFalls(xJobCode, Optional xNewValue, Optional xGridType)
Dim rsHRJob As New ADODB.Recordset
Dim xBaseAmt
Dim xDHrs
Dim xPHrs
Dim xSalCode, xUnion
Dim xPayPeriod

    rsHRJob.Open "SELECT JB_CODE, JB_ORG, JB_DHRS, JB_S1, JB_SALCD FROM HRJOB WHERE JB_CODE = '" & xJobCode & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRJob.EOF Then
        If rsHRJob("JB_SALCD") = "H" Then
            If IsMissing(xNewValue) Then
                Get_Hourly_Rate_NiagaraFalls = rsHRJob("JB_S1")
            Else
                Get_Hourly_Rate_NiagaraFalls = xNewValue
            End If
        ElseIf rsHRJob("JB_SALCD") = "A" And (Not IsNull(rsHRJob("JB_S1")) And rsHRJob("JB_S1") <> "") Then
            'Get Pay Period from Payroll Matrix using the Position Union from Position Master
            xPayPeriod = Get_Pay_Period(rsHRJob("JB_ORG"))
            If xPayPeriod = "" Or xPayPeriod = 0 Then
                xPayPeriod = 1
            End If
            
            If IsNull(rsHRJob("JB_DHRS")) Or rsHRJob("JB_DHRS") = "" Then
                'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                If IsMissing(xNewValue) Then
                    'Get_Hourly_Rate_NiagaraFalls = Round((rsHrJob("JB_S1") / xPayPeriod) / (1 * 5), 4)
                    Get_Hourly_Rate_NiagaraFalls = Round((rsHRJob("JB_S1") / xPayPeriod) / (1), 4)
                Else
                    'Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / (1 * 5), 4)
                    Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / (1), 4)
                End If
            Else
                'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                If IsMissing(xNewValue) Then
                    'Get_Hourly_Rate_NiagaraFalls = Round((rsHrJob("JB_S1") / xPayPeriod) / (rsHrJob("JB_DHRS") * 5), 4)
                    Get_Hourly_Rate_NiagaraFalls = Round((rsHRJob("JB_S1") / xPayPeriod) / rsHRJob("JB_DHRS"), 4)
                Else
                    'Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / (rsHrJob("JB_DHRS") * 5), 4)
                    Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / rsHRJob("JB_DHRS"), 4)
                End If
            End If
        Else
            If IsNull(rsHRJob("JB_S1")) Or rsHRJob("JB_S1") = "" Then
                If IsMissing(xNewValue) Then
                    xBaseAmt = frmMPOSITIONS.medPayScale(1).Text
                Else
                    xBaseAmt = xNewValue
                End If
                xDHrs = frmMPOSITIONS.medHours.Text
                xSalCode = frmMPOSITIONS.lblSalCode.Caption
                xUnion = frmMPOSITIONS.clpCode(3).Text
                'Get Pay Period from Payroll Matrix using the Position Union from Position Master
                xPayPeriod = Get_Pay_Period(xUnion)
                If xPayPeriod = "" Or xPayPeriod = 0 Then
                    xPayPeriod = 1
                End If
                
                If xSalCode = "A" Then
                    'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                    If xBaseAmt <> "" And (xDHrs <> "" Or xDHrs <> 0) Then
                        'Get_Hourly_Rate_NiagaraFalls = Round((xBaseAmt / xPHrs) / (xDHRS * 5), 4)
                        Get_Hourly_Rate_NiagaraFalls = Round((xBaseAmt / xPHrs) / xDHrs, 4)
                    Else
                        Get_Hourly_Rate_NiagaraFalls = 0
                    End If
                Else
                    If IsMissing(xNewValue) Then
                        Get_Hourly_Rate_NiagaraFalls = xBaseAmt
                    Else
                        Get_Hourly_Rate_NiagaraFalls = xNewValue
                    End If
                End If
            Else
                Get_Hourly_Rate_NiagaraFalls = 0
            End If
        End If
    Else
        If IsMissing(xNewValue) Then
            Get_Hourly_Rate_NiagaraFalls = 0
        Else
            If Not IsMissing(xGridType) Then
                If xGridType = "H" Then
                    Get_Hourly_Rate_NiagaraFalls = xNewValue
                ElseIf xGridType = "A" And (Not IsNull(xNewValue) And xNewValue <> "") Then
                    'Get Pay Period from Payroll Matrix using the Position Union from Position Master
                    xPayPeriod = Get_Pay_Period(frmMPOSITIONS.clpCode(3).Text)
                    xDHrs = frmMPOSITIONS.medHours.Text
                    If xPayPeriod = "" Or xPayPeriod = 0 Then
                        xPayPeriod = 1
                    End If
                    
                    If IsNull(xDHrs) Or xDHrs = "" Then
                        'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                        'Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / (1 * 5), 4)
                        Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / 1, 4)
                    Else
                        'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                        'Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / (rsHrJob("JB_DHRS") * 5), 4)
                        Get_Hourly_Rate_NiagaraFalls = Round((xNewValue / xPayPeriod) / xDHrs, 4)
                    End If
                Else
                    xBaseAmt = xNewValue
                    xDHrs = frmMPOSITIONS.medHours.Text
                    xSalCode = frmMPOSITIONS.lblSalCode.Caption
                    xUnion = frmMPOSITIONS.clpCode(3).Text
                    
                    'Get Pay Period from Payroll Matrix using the Position Union from Position Master
                    xPayPeriod = Get_Pay_Period(xUnion)
                    If xPayPeriod = "" Or xPayPeriod = 0 Then
                        xPayPeriod = 1
                    End If
                    
                    If xSalCode = "A" Then
                        'Hemu - Ticket #16071 - Annual Salary / Pay Periods Per Year / Hours Per Pay
                        If xBaseAmt <> "" And (xDHrs <> "" Or xDHrs <> 0) Then
                            'Get_Hourly_Rate_NiagaraFalls = Round((xBaseAmt / xPHrs) / (xDHRS * 5), 4)
                            Get_Hourly_Rate_NiagaraFalls = Round((xBaseAmt / xPHrs) / xDHrs, 4)
                        Else
                            Get_Hourly_Rate_NiagaraFalls = 0
                        End If
                    Else
                        Get_Hourly_Rate_NiagaraFalls = xNewValue
                    End If
                End If
            End If
        End If
    End If
    rsHRJob.Close

End Function

Public Function GetDeptName(xDeptno, xField)
Dim rsDEPT As New ADODB.Recordset
Dim SQLQ As String

    SQLQ = "SELECT DF_NBR," & xField & " FROM HRDEPT WHERE DF_NBR = '" & xDeptno & "'"
    rsDEPT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDEPT.EOF Then
        GetDeptName = IIf(IsNull(rsDEPT(xField)), "", rsDEPT(xField))
    Else
        GetDeptName = ""
    End If
    rsDEPT.Close
    Set rsDEPT = Nothing
End Function

Public Function Get_Pay_Period(xUnion)
Dim rsHRMatrix As New ADODB.Recordset
Dim SQLQ As String

    SQLQ = "SELECT M_DEFTYPE, M_CODE, M_CONVERT1 FROM HRMATRIX WHERE M_CODE ='" & xUnion & "' AND M_DEFTYPE = 'UNIO'"
    rsHRMatrix.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRMatrix.EOF Then
        If IsNull(rsHRMatrix("M_CONVERT1")) Or rsHRMatrix("M_CONVERT1") = "" Then
            Get_Pay_Period = 1
        Else
            Get_Pay_Period = Val(rsHRMatrix("M_CONVERT1"))
        End If
    End If
    rsHRMatrix.Close
    Set rsHRMatrix = Nothing
End Function

Private Function PayrollMatrix(MType As String, mcode, Optional ConvertFld As String = "M_CONVERT1") As String
    Dim rsPayrollMatrix As New ADODB.Recordset
    rsPayrollMatrix.Open "SELECT " & ConvertFld & " FROM HRMATRIX WHERE M_TYPE='" & MType & "' AND M_CODEDEPT='" & mcode & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    PayrollMatrix = mcode
    If Not rsPayrollMatrix.EOF Then
        If Not IsNull(rsPayrollMatrix(ConvertFld)) Then
            PayrollMatrix = rsPayrollMatrix(ConvertFld)
        End If
    End If
End Function

Public Sub Update_HR_Vadim_Sy_Interface(xBatchID, xEmpnbr, xEffDate, xSalary, xJob, xType, xTransDt, xNotes)
    Dim rsHRVadimInt As New ADODB.Recordset
    Dim SQLQ As String
    
    If xType = "A" Or xType = "M" Then  'M = other salary info. update comes in as "M"
        'Add a new entry in the HR_VADIM_SY_INTERFACE table
        SQLQ = "SELECT * FROM HR_VADIM_SY_INTERFACE WHERE 1=2"
        rsHRVadimInt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        rsHRVadimInt.AddNew
        rsHRVadimInt("VS_EMPNBR") = xEmpnbr
        rsHRVadimInt("VS_EFF_DATE") = xEffDate
        rsHRVadimInt("VS_SY_BATCH_ID") = xBatchID
        rsHRVadimInt("VS_SALARY") = xSalary
        rsHRVadimInt("VS_JOB") = xJob
        rsHRVadimInt("VS_PROC_TYPE") = xType
        rsHRVadimInt("VS_TRANS_DATE") = xTransDt
        rsHRVadimInt("VS_NOTES") = xNotes
        rsHRVadimInt.Update
        
        rsHRVadimInt.Close
        
    Else    'if xType = D or xType = C (existing future salary change)
        'Delete existing entry from HR_VADIM_SY_INTERFACE table after retrieving the Batch ID
        SQLQ = "SELECT * FROM HR_VADIM_SY_INTERFACE WHERE "
        SQLQ = SQLQ & " VS_EMPNBR = " & xEmpnbr
        SQLQ = SQLQ & " AND VS_EFF_DATE = " & Date_SQL(xEffDate)
        SQLQ = SQLQ & " AND VS_SALARY = " & xSalary
        SQLQ = SQLQ & " AND VS_JOB = '" & xJob & "'"
        rsHRVadimInt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRVadimInt.EOF Then
            rsHRVadimInt.MoveFirst
            
            Do While Not rsHRVadimInt.EOF
                'Call procedure to delete the Sy_Interface records in the Vadim tables using the Batch Id retrieved.
                Call Delete_SY_INTERFACE_Records(rsHRVadimInt("VS_SY_BATCH_ID"))
                
                'Delete the entry from HR_VADIM_SY_INTERFACE
                rsHRVadimInt.Delete
                
                rsHRVadimInt.MoveNext
            Loop
        End If
        rsHRVadimInt.Close
    End If

    Set rsHRVadimInt = Nothing
End Sub

Private Sub Delete_SY_INTERFACE_Records(xBatchID)
        
    'Delete the records from Vadim's Sy_Interface and Sy_Interface_Batch tables for the Batch ID provided.
    gdbPayroll.BeginTrans
    gdbPayroll.Execute "DELETE FROM SY_INTERFACE WHERE SY_BATCH_ID=" & xBatchID
    gdbPayroll.Execute "DELETE FROM SY_INTERFACE_BATCH WHERE SY_BATCH_ID=" & xBatchID
    gdbPayroll.CommitTrans
    
End Sub

Public Sub Compute_Salary_Vadim_Based(xEmpnbr, xSalCD, xSalary, xPHrs, xWHRS, ByRef newSalary, ByRef newRate)
Dim X
Dim UptType
'Dim newSalary
'Dim newRate

newRate = 0
newSalary = 0
If xSalCD = "H" Then
    'City of Niagara Falls - Round the Hourly Rate to 2 Decimal places - Hourly to Hourly Rate conversion
    If glbCompSerial = "S/N - 2276W" Then
        newRate = Round(xSalary, 4)  'round to 4 decimal places instead of 2
        newSalary = Round(newRate * Val(frmESALARY.txtWHRS.Text), 4) 'round to 4 decimal places
    Else
        newRate = xSalary
        newSalary = Round2DEC(newRate * xPHrs)
    End If
    
ElseIf xSalCD = "A" Then
    'City of Niagara Falls - Special formula to calculate Hourly Rate - xWHRS actually contains Hours per Day
    If glbCompSerial = "S/N - 2276W" Then
        'Round the Hourly Rate to 4 Decimal places - Salary to Hourly Rate conversion
        If xWHRS <> 0 Then newRate = Round((xSalary / xPHrs) / (xWHRS * 5), 4)
        newSalary = Round(newRate * Val(frmESALARY.txtWHRS.Text), 4) 'round to 4 decimal places
    Else
        If xWHRS <> 0 Then newRate = Round2DEC((xSalary / 52) / xWHRS)
        newSalary = Round2DEC(((xSalary / 52) / xWHRS) * xPHrs)
    End If
ElseIf xSalCD = "M" Then
    If xWHRS <> 0 Then newRate = Round2DEC(((xSalary * 12) / 52) / xWHRS)
    newSalary = Round2DEC(newRate * xPHrs)
End If

End Sub

Public Sub Update_VadimDB_HR_EMP_HISTORY(xPayID, UptDate, oldLevel, newLevel, xJob, UpdType, Optional xEndDate)
Dim xBatchID
Dim VDClause As String

    
    xBatchID = AddBatchVadim(UpdType, UptDate)
    If UpdType = "M" Or UpdType = "D" Then
        VDClause = "EMP_HIST_DATE='" & Format(UptDate, "YYYY/MM/DD") & "'"
    Else
        VDClause = ""
    End If

    If UpdType = "A" Then
        Call PassDataToVadim(xBatchID, "HR_EMP_HIST.EMP_NUM", xPayID, "", xPayID)
        Call PassDataToVadim(xBatchID, "HR_EMP_HIST.EMP_HIST_DATE", xPayID, "", UptDate)
        'Ticket #16160
        'Call PassDataToVadim(xBatchID, "HR_EMP_HIST.EMP_HIST_VALUE", xPayid, "", newLevel)
        Call PassDataToVadim(xBatchID, "HR_EMP_HIST.RATE_LEVEL", xPayID, "", newLevel)
        Call PassDataToVadim(xBatchID, "HR_EMP_HIST.EMP_HIST_CODE_VAL", xPayID, "", xJob)
    ElseIf UpdType = "M" Then
        If Not IsMissing(xEndDate) Then
            'Pass the End Date
            Call PassDataToVadim(xBatchID, "HR_EMP_HIST.EMP_HIST_END_DATE", xPayID, "", xEndDate, VDClause)
        Else
            'Ticket #16160
            'Call PassDataToVadim(xBatchID, "HR_EMP_HIST.EMP_HIST_VALUE", xPayid, oldLevel, newLevel, VDClause)
            Call PassDataToVadim(xBatchID, "HR_EMP_HIST.RATE_LEVEL", xPayID, oldLevel, newLevel, VDClause)
        End If
    ElseIf UpdType = "D" Then
        Call PassDataToVadim(xBatchID, "HR_EMP_HIST.EMP_NUM", xPayID, xPayID, "", VDClause)
    End If
    
    'Update_VadimDB_HR_EMP_HISTORY(xPayid, UpdDate, oldRate, newRate, GetJHData(xPayid, "JH_JOB", ""), "M")
    'Call PassDataToVadim(xBatchID, VDTableFields(Y), xPayid, oldValue, NewValue, VDClause)

    Call CloseBatchVadim(xBatchID)
    
End Sub

