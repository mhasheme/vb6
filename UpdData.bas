Attribute VB_Name = "UpdData"
Option Explicit
Global glbFlag_BenefitForSalDEPN
Global glbSalaryEDate
Type BenefitCost
    Salary As Double
    Type As String
End Type

Public Enum DependentRelationship
    Spouse
    Children
    Other
    Employee
End Enum
'    Child
'    Aunt
'    Brother
'    [Common Law]
'    Couple
'    Daughter
'    Estate
'    [Ex-Spouse]
'    Father
'    Fiancee
'    Fiance
'    Husband
'    Mother
'    Other
'    Parents
'    Sister
'    Son
'    Uncle
'    Wife

Function CGLUpdate(updEMPID, Optional newEMPID, Optional DeleteRecord)
Dim cnCGL As New ADODB.Connection
Dim rsCGL As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim SQLQ
Dim xPCODE, xPHONE, xSHIFT
Dim xEmpNbr
On Error GoTo Err_CGLUpdate
'Exit Function ' do not run this procedure until prove from CGL.

If glbCompSerial <> "S/N - 2349W" Then Exit Function
cnCGL.Mode = adModeReadWrite

cnCGL.Open "Provider=OraOLEDB.Oracle.1;Password=" & SQLUserPassword & ";Persist Security Info=True;User ID=" & SQLUserName & ";Data Source=" & SQLServerName & ""

rsCGL.Open "SELECT * FROM EMPLOYEE WHERE ID='" & updEMPID & "'", cnCGL, adOpenStatic, adLockPessimistic
If Not IsMissing(DeleteRecord) Then
    If Not rsCGL.EOF Then
        rsCGL.Delete
        rsCGL.Close
    End If
    Exit Function
    
End If

xEmpNbr = IIf(IsNumeric(newEMPID), newEMPID, updEMPID)
SQLQ = "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME,ED_EMP,"
SQLQ = SQLQ & " ED_MIDNAME,ED_ADDR1,ED_ADDR2,ED_DOH,ED_CITY,"
SQLQ = SQLQ & " ED_PROV,ED_PCODE,ED_COUNTRY,ED_PHONE,ED_SHIFT "
SQLQ = SQLQ & " FROM HREMP WHERE ED_EMPNBR=" & xEmpNbr

rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
SQLQ = "SELECT SH_EMPNBR,SH_SALARY,SH_SALCD,SH_EDATE "
SQLQ = SQLQ & " FROM HR_SALARY_HISTORY "
SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & xEmpNbr
rsSal.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If rsEmp.EOF Or rsSal.EOF Then
    If Not rsCGL.EOF Then
        rsCGL!Active = "N"
        rsCGL.Update
    End If
Else

    If rsCGL.EOF Then
        rsCGL.AddNew
        rsCGL!ID = xEmpNbr
    End If
    rsCGL!LAST_NAME = rsEmp!ED_SURNAME
    rsCGL!FIRST_NAME = rsEmp!ED_FNAME
    rsCGL!MIDDLE_INITIAL = Left(Format(rsEmp!ED_MIDNAME, "@"), 1)
    rsCGL!ADDR_1 = Format(rsEmp!ED_ADDR1, "@") & " " & Format(rsEmp!ED_ADDR2, "@")
    If IsDate(rsEmp!ED_DOH) Then rsCGL!ADDR_2 = Format(rsEmp!ED_DOH, "MM/DD/YYYY")
    If IsDate(rsSal!SH_EDATE) Then rsCGL!ADDR_3 = Format(rsSal!SH_EDATE, "MM/DD/YYYY")
    rsCGL!City = rsEmp!ED_CITY
    rsCGL!State = rsEmp!ED_PROV & "T"
    
    xPCODE = Format(rsEmp!ED_PCODE, "@")
    xPCODE = Replace(xPCODE, " ", "")
    xPCODE = Left(xPCODE & String(6, " "), 6)
    xPCODE = Left(xPCODE, 3) & " " & Mid(xPCODE, 4)
    rsCGL!ZIPCODE = Trim(xPCODE)
    
    rsCGL!Country = rsEmp!ED_Country
    
    xPHONE = rsEmp!ED_PHONE
    xPHONE = Replace(xPHONE, " ", "")
    xPHONE = Replace(xPHONE, "-", "")
    xPHONE = Replace(xPHONE, "(", "")
    xPHONE = Replace(xPHONE, ")", "")
    xPHONE = Left(xPHONE & String(10, " "), 10)
    xPHONE = "(" & Left(xPHONE, 3) & ")" & Mid(xPHONE, 4, 3) & "-" & Mid(xPHONE, 7)
    rsCGL!Phone = Trim(xPHONE)
    
    If rsSal!SH_SALCD = "A" Then
        rsCGL!BASE_PAY_RATE = Null
    Else
        rsCGL!BASE_PAY_RATE = rsSal!SH_SALARY
    End If
    rsCGL!Type = IIf(rsSal!SH_SALCD = "H", "H", "S")
    rsCGL!DEPARTMENT_ID = "DEFAULT"
    rsCGL!EARNING_CODE_ID = "DEF"
    
    rsCGL!Active = IIf(Format(rsEmp!ED_EMP, "@") = "1", "Y", "N")
    xSHIFT = Trim(Format(rsEmp!ed_shift, "@"))
    rsCGL!SHIFT_ID = IIf(xSHIFT = "1", "DAY", IIf(xSHIFT = "2", "AFT", IIf(xSHIFT = "3", "NIGHT", "DAY")))
    rsCGL.Update

End If
rsEmp.Close
rsSal.Close
rsCGL.Close
cnCGL.Close
Exit Function
Err_CGLUpdate:
    If InStr(Err.Description, "ORA") Then
        MsgBox Err.Description & vbNewLine & vbNewLine & "Please check Custom Features - Visual Quality System Data Source Settings"
        Exit Function
    End If
    glbFrmCaption$ = "Visual Quality System Interface"
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CGLUpdate", "Employee", "UPDATE")
End Function
Public Sub updCompPlan(updEMPID, difSalary, Edate)
Dim cnCP As New ADODB.Connection
Dim xEDate, SQLQ
Dim rsCP As New ADODB.Recordset

cnCP.Open Replace(UCase(glbAdoIHRDB), "IHR001.MDB", "SN2291.MDB")

SQLQ = SQLQ & "SELECT * FROM COMPENSATION_PLAN "
SQLQ = SQLQ & "WHERE CP_EMPNBR = " & updEMPID
SQLQ = SQLQ & " ORDER BY CP_FDATE DESC"

rsCP.Open SQLQ, cnCP, adOpenStatic, adLockPessimistic
If rsCP.EOF Then Exit Sub
rsCP.MoveFirst
Do Until rsCP.EOF
    If rsCP("CP_FDATE") <= CVDate(Edate) And rsCP("CP_TDATE") > CVDate(Edate) Then
        rsCP("CP_ADJSALARY") = difSalary + rsCP("CP_ADJSALARY") '- xSalary(0)
        rsCP("CP_TARGINCOME") = difSalary + rsCP("CP_TARGINCOME")
        rsCP.Update
    End If
    rsCP.MoveNext
Loop
rsCP.Close
Exit Sub


End Sub

Function ElginBenefit(updEMPID, xSalary)
Dim rsTemp As New ADODB.Recordset, rsEmp As New ADODB.Recordset
Dim rsBenefit As New ADODB.Recordset
Dim xED_PT, xED_ORG, xJB_GRPVD, xCovAmt, IfElginLife As Boolean
Dim Flag1, Flag2, Flag3, txtCode
Dim xDate, xYY, xMM, xDD, txtWaitPeriod
Dim xSalFactor, xRndFactor, xMaxCover, xMinCover
Dim xCovAmount, xPer
Dim xRound
Dim xUNITCOST, xTCost, SQLQ
    ElginBenefit = False
    SQLQ = "SELECT HREMP.ED_EMPNBR, HREMP.ED_PT, HREMP.ED_ORG, HREMP.ED_DOH, HRJOB.JB_GRPCD "
    SQLQ = SQLQ & "FROM (HREMP LEFT JOIN HR_JOB_HISTORY ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR) LEFT JOIN HRJOB ON "
    SQLQ = SQLQ & "HR_JOB_HISTORY.JH_JOB = HRJOB.JB_CODE "
    SQLQ = SQLQ & "WHERE (ED_EMPNBR = " & updEMPID & " And (HR_JOB_HISTORY.JH_CURRENT) <> 0)"
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF And rsTemp.BOF Then
        xED_PT = ""
        xED_ORG = ""
        xJB_GRPVD = ""
    Else
        xED_PT = rsTemp("ED_PT")
        xED_ORG = rsTemp("ED_ORG")
        xJB_GRPVD = rsTemp("JB_GRPCD")
    End If
    rsTemp.Close
    
    If xED_PT <> "FT" Then Exit Function

    If Not (xED_ORG = "3" Or xED_ORG = "1" Or xED_ORG = "5") Then Exit Function

    'SQLQ = "SELECT * FROM HRBENFT WHERE (BF_EMPNBR = " & updEMPID & ")"
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & updEMPID & " AND BF_BCODE IN ('LIFE','AD&D')"
    rsBenefit.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsBenefit.EOF And rsBenefit.BOF Then
        Exit Function
    End If
    Do While Not rsBenefit.EOF
        txtCode = rsBenefit("BF_BCODE")
        ' danielk - 04/15/2003 - Ticket #4020
        xUNITCOST = rsBenefit("BF_UNITCOST")
        Flag1 = False
        Flag2 = False
        Flag3 = False
    
        If UCase(txtCode) = "LIFE" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
            Flag1 = True
        End If

        If UCase(txtCode) = "LIFE" And xED_PT = "FT" And xED_ORG = "1" Then
            Flag1 = True
        End If

        If UCase(txtCode) = "AD&D" And xED_PT = "FT" And xED_ORG = "5" And (xJB_GRPVD = "ADMN" Or xJB_GRPVD = "MGMT") Then
            Flag2 = True
        End If
        If UCase(txtCode) = "AD&D" And xED_PT = "FT" And xED_ORG = "3" And (xJB_GRPVD = "ONA") Then
            Flag3 = True
        End If
    
        If Flag1 Or Flag2 Or Flag3 Then
            
'            txtWaitPeriod = 90
'            rsEMP.Open "SELECT ED_DOH FROM HREMP WHERE ED_EMPNBR=" & updEMPID, gdbAdoIhr001, adOpenStatic
'            If Not rsEMP.EOF Then
'                If IsDate(rsEMP("ED_DOH")) Then
'                        xDATE = DateAdd("d", txtWaitPeriod, rsEMP("ED_DOH"))
'                        xDD = Day(CVDate(xDATE))
'                        If xDD > 15 Then
'                            xDATE = DateAdd("d", -(xDD - 1), CVDate(xDATE))
'                            xDATE = DateAdd("m", 1, CVDate(xDATE))
'                        End If
'                End If
'            End If
'            rsEMP.Close
'            rsBenefit("BF_WaitPeriod") = txtWaitPeriod
'            If IsDate(xDATE) Then
'                rsBenefit("BF_EDATE") = CVDate(xDATE)
'            End If

            rsBenefit("BF_SALARYDEPENDANT") = "Y"
'            rsBenefit("BF_COVER") = "Y"
'            xMinCover = 0
'            rsBenefit("BF_MINIMUM") = xMinCover
            xMinCover = rsBenefit("BF_MINIMUM")
            'Comment By Franks Jun 3,2002
            'xMaxCover = 200000
            'rsBenefit("BF_MAXIMUM") = xMaxCover
            xMaxCover = rsBenefit("BF_MAXIMUM")
'            xSalFactor = 2
'            rsBenefit("BF_FACTOR") = xSalFactor
            xSalFactor = rsBenefit("BF_FACTOR")
            
'            rsBenefit("BF_NEXTNEAREST") = "N"
            xRound = rsBenefit("BF_NEXTNEAREST")
'            xRndFactor = 1000
'            rsBenefit("BF_ROUND") = xRndFactor
            xRndFactor = rsBenefit("BF_ROUND")
'            xUNITCOST = 0.372
'            rsBenefit("BF_UNITCOST") = xUNITCOST
'            xPER = 1000
'            rsBenefit("BF_PER") = xPER
            xPer = rsBenefit("BF_PER")
'            rsBenefit("BF_PCC") = 1
'            rsBenefit("BF_PCE") = 0
'            rsBenefit("BF_PREMIUM") = "P"
'            rsBenefit("BF_TAXBEN") = "Y"

            xCovAmount = xSalary * xSalFactor 'xCovAmt * xSalFactor
            If xMinCover <> 0 And xCovAmount < xMinCover Then xCovAmount = xMinCover
            If xMaxCover <> 0 And xCovAmount > xMaxCover Then xCovAmount = xMaxCover
            If xRndFactor = 0 Then xRndFactor = 0.01
'            xCovAmount = Round(xCovAmount / xRndFactor + 0.49) * xRndFactor
            xCovAmount = Round(xCovAmount / xRndFactor + IIf(xRound = "N", 0.5, 0)) * xRndFactor
            rsBenefit("BF_AMT") = xCovAmount
            
'            If Flag2 Or Flag3 Then
                ' danielk - 12/17/2002 - commented out the tax addition below, they were adding tax on
                ' in their custom report as well as here, causing the numbers to be too high.  Doesn't
                ' know why they asked for it to be added here.
                ' Ticket #3655
'                If xPER > 0 Then xTCost = (xCovAmount / xPER * xUNITCOST) * 12 '* 1.08
'            Else
'                If xPER > 0 Then xTCost = (xCovAmount / xPER) * xUNITCOST * 12
'            End If
            
            If xPer > 0 Then xTCost = (xCovAmount / xPer) * xUNITCOST
            rsBenefit("BF_TCOST") = xTCost
'            rsBenefit("BF_ECOST") = xTCost * 0
'            rsBenefit("BF_CCOST") = xTCost * 1
            rsBenefit("BF_ECOST") = xTCost * rsBenefit("BF_PCE")
            rsBenefit("BF_CCOST") = xTCost * rsBenefit("BF_PCC")
'            rsBenefit("BF_MTHECOST") = 0
            rsBenefit("BF_MTHCCOST") = Round(rsBenefit("BF_CCOST") / 12, 2)
            rsBenefit.Update

        End If
    
        rsBenefit.MoveNext
    Loop
    rsBenefit.Close
    ElginBenefit = True
    
End Function

Public Sub updBenefitForSurreyPlace(updEMPID)
Dim SQLQ
Dim FTE
Dim USDate
Dim rsBENF As New ADODB.Recordset
Dim NomalCCost
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xPT, xDiv

If glbCompSerial <> "S/N - 2347W" Then Exit Sub
'Jerry asked to remove the DENT and SHHP
'rsBENF.Open "SELECT * FROM HRBENFT WHERE (BF_BCODE = 'DENT' OR BF_BCODE = 'SHHP' OR BF_BCODE = 'VH') AND BF_EMPNBR=" & updEMPID, gdbAdoIhr001, adOpenStatic, adLockPessimistic
rsBENF.Open "SELECT * FROM HRBENFT WHERE BF_BCODE = 'VH' AND BF_EMPNBR=" & updEMPID, gdbAdoIhr001, adOpenStatic, adLockPessimistic
FTE = GetJHData(updEMPID, "JH_FTENUM", 1)
Do Until rsBENF.EOF
    NomalCCost = rsBENF("BF_TCOST") * rsBENF("BF_PCC")
    rsBENF("BF_CCOST") = FTE * NomalCCost
    rsBENF("BF_ECOST") = rsBENF("BF_TCOST") * rsBENF("BF_PCE") + (1 - FTE) * NomalCCost
    rsBENF("BF_MTHCCOST") = rsBENF("BF_CCOST") / 12
    rsBENF("BF_MTHECOST") = rsBENF("BF_ECOST") / 12
    rsBENF.Update
    
    'Update Audit also otherwise the changes are not exported to ADP -----------------------
    rsTA.Open "SELECT * FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
    rsTA.AddNew
    
    rsTB.Open "SELECT ED_PT,ED_DIV,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        'xPT = rsTB("ED_PT")
        'xDiv = rsTB("ED_DIV")
        If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
        If IsNull(rsTB("ED_DIV")) Then xDiv = "" Else xDiv = rsTB("ED_DIV")
        If Not IsNull(rsTB("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsTB("ED_PAYROLL_ID")
    Else
        xPT = ""
        xDiv = ""
    End If
    
    rsTA("AU_EMPNBR") = updEMPID
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDiv
    rsTA("AU_BCODE") = "VH"
    rsTA("AU_MTHCCOST") = rsBENF("BF_MTHCCOST")
    rsTA("AU_MTHECOST") = rsBENF("BF_MTHECOST")
    rsTA("AU_COMPNO") = "001"
    
    
    If CVDate(Format(rsBENF("BF_EDATE"), "SHORT DATE")) >= CVDate(Format(Now, "SHORT DATE")) Then
      rsTA("AU_LDATE") = Format(rsBENF("BF_EDATE"), "SHORT DATE")
    Else
      rsTA("AU_LDATE") = Date
    End If
    
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "M"
    rsTA.Update
    rsTB.Close
    rsTA.Close
    
    '---------------------------------------------------------------------------------------
    
    rsBENF.MoveNext
Loop
rsBENF.Close
End Sub

Public Sub updBenefitForSalDEPN(updEMPID, Optional FromBenefitGroup, Optional xTranDate)
Dim SQLQ, SQLW, xSalary, xSalaryElgin
Dim FlagElgin As Boolean
Dim rsBF As New ADODB.Recordset
Dim SalaryDependantFind
Dim zSalary
Dim CostINFO As BenefitCost
Dim xOmersFormula As Boolean
Dim xIfDatChg As Boolean 'Ticket #23729 Franks 05/10/2013
Dim xPreVal 'Ticket #23729 Franks 05/13/2013
Dim rsBenWrk As New ADODB.Recordset 'Ticket #23729 Franks 05/13/2013

If glbWFC Then 'Ticket #23247 Franks 04/26/2013
    If IsWFCUSBenEmp(updEMPID) Then
        If Not IsMissing(xTranDate) Then
            Call WFC_UptUSBenByEmp(updEMPID, CVDate(xTranDate), 0, "Y", "Y")
        Else
            Call WFC_UptUSBenByEmp(updEMPID, CVDate(Date), 0, "Y", "Y")
        End If
        Exit Sub
    End If
End If

xOmersFormula = OMER_UseCostTable 'Ticket #20872 Franks 09/27/2011

SQLW = "BF_SALARYDEPENDANT='Y' AND BF_EMPNBR=" & updEMPID
If Not IsMissing(FromBenefitGroup) Then If FromBenefitGroup Then SQLW = SQLW & " AND BF_LUSER='999999998'"

SQLQ = "SELECT BF_EMPNBR,BF_GROUP,BF_BCODE,BF_BENE_ID,BF_EDATE FROM HRBENFT WHERE " & SQLW

rsBF.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
SalaryDependantFind = False
'Comment by Frank, Ticket #15270, this Begin caused a problem of incorrect BF_AMT calculation
'gdbAdoIhr001.BeginTrans

Do Until rsBF.EOF
    SalaryDependantFind = True
    glbFlag_BenefitForSalDEPN = True
    xSalary = CrtSalary(updEMPID)
    CostINFO = CrtBeneCost(updEMPID, xSalary, rsBF("BF_GROUP").Value, rsBF("BF_BCODE").Value)
    xSalary = CostINFO.Salary
    
    ''Ticket #28969 - Moving this logic after the BF_LDATE update because, for Vadim Integration, a change to a Benefit record
    ''triggers an update to Vadim with the invalid BF_LDATE as the Update Date, esp. when it's a future dated update.
    'If glbCompSerial = "S/N - 2205W" Then 'Crown Investment Corp. - Ticket #14090
    '    If rsBF("BF_BCODE") = "GLFE" Then
    '        SQLQ = "UPDATE HRBENFT SET BF_AMT=(" & xSalary & " * BF_FACTOR) - 25000 WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
    '    Else
    '        SQLQ = "UPDATE HRBENFT SET BF_AMT=" & xSalary & " * BF_FACTOR WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
    '    End If
    'Else
    '    SQLQ = "UPDATE HRBENFT SET BF_AMT=" & xSalary & " * BF_FACTOR WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
    'End If
    'gdbAdoIhr001.Execute SQLQ
    
    '------------------------------------------------------------------------------------------------------------------------
    'Ticket #27694 Hemu 10/27/2015 - Commented and Fixed the logic code below. For New Hires with future Benefit Effective Date
    'because of the waiting period, the BF_LDATE was getting updated with today's date. This was causing the Benefits for new
    'hires to be exported right away. Changed from checking with Salary Effective Date (which will normally be same as the hire
    'date for new hires) to Benefit Effective Date which already includes the waiting period. This fixed the issue but
    'then came across another scenario where an existing employee gets Salary Update for future date where the Salary
    'Dependent Benefits should be updated as well. In this scenario the Benefit Effective Date would most likely be an
    'older date so checking only with Benefit Effective Date will not work in this screnario as it will export the Salary
    'Dependent Benefit changes right away when the Salary is still not effective yet. So added another checking with
    'Salary Effective Date as well. So whichever date (Salary or Benefit Effective) is more future, that date will be used
    'to update the BF_LDATE hence that will be the export date for that Benefit.
    
'    'Ticket #26928 Franks 04/09/2015 - if Salary Effective Date is future date,
'    'then the AU_LDATE will be future date too, it is Salary Effective Date(use BF_LDATE to update AU_LDATE)
'    If IsDate(glbSalaryEDate) Then
'        If CVDate(glbSalaryEDate) > CVDate(Date) Then
'            SQLQ = "UPDATE HRBENFT SET BF_LDATE=" & Date_SQL(glbSalaryEDate) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
'        Else
'            SQLQ = "UPDATE HRBENFT SET BF_LDATE=" & Date_SQL(Date) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
'        End If
'        gdbAdoIhr001.Execute SQLQ
'    End If
    
    If IsDate(rsBF("BF_EDATE")) And IsDate(glbSalaryEDate) Then
        'If CVDate(glbSalaryEDate) > CVDate(Date) Then
        If CVDate(rsBF("BF_EDATE")) > CVDate(glbSalaryEDate) Then
            If CVDate(rsBF("BF_EDATE")) > CVDate(Date) Then
                'SQLQ = "UPDATE HRBENFT SET BF_LDATE=" & Date_SQL(glbSalaryEDate) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
                SQLQ = "UPDATE HRBENFT SET BF_LDATE=" & Date_SQL(rsBF("BF_EDATE")) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
            Else
                SQLQ = "UPDATE HRBENFT SET BF_LDATE=" & Date_SQL(Date) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
            End If
        Else
            If CVDate(glbSalaryEDate) > CVDate(Date) Then
                SQLQ = "UPDATE HRBENFT SET BF_LDATE=" & Date_SQL(glbSalaryEDate) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
            Else
                SQLQ = "UPDATE HRBENFT SET BF_LDATE=" & Date_SQL(Date) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
            End If
        End If
        gdbAdoIhr001.Execute SQLQ
    End If
    '------------------------------------------------------------------------------------------------------------------------
    
    'Ticket #19145 - Disable this custom code
    'Ticket #13287 Frank 07/05/2007
    'Date salary became effective if it's a salary dependent benefit that was changed during a change of salary
    'If glbWFC Then
    '    If Len(glbSalaryEDate) > 0 Then
    '        SQLQ = "UPDATE HRBENFT SET BF_EDATE=" & Date_SQL(glbSalaryEDate) & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
    '        gdbAdoIhr001.Execute SQLQ
    '    End If
    'Else
        ''The Walter Fedy Partnership - Ticket #15298
        'If (glbCompSerial = "S/N - 2386W") Then
        'Ticket #28065 Franks 01/29/2016 - make this for all
            'Ticket #28969 - This is causing as issue once it was opened for all, as the above logic where it
            'assigns the right BF_LDATE is getting overwritten by this one when there glbSalaryEDate. I have
            'added condition to only update if glbSalaryEDate is not a date.
            If IsDate(rsBF("BF_EDATE")) And Not IsDate(glbSalaryEDate) Then
                SQLQ = "UPDATE HRBENFT SET BF_LDATE=BF_EDATE WHERE BF_EDATE >" & Date_SQL(Now) & " AND " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
                gdbAdoIhr001.Execute SQLQ
            End If
        'End If
    'End If
    
    'Ticket #28969 - This logic has been moved from the top to after the BF_LDATE update because, for Vadim Integration,
    'a change to a Benefit record triggers an update to Vadim with the invalid BF_LDATE as the Update Date, esp. when
    'it's a future dated update. This will make sure the Benefit transfer is done on the right day as the BF_LDATE has been
    'updated with the right Update/transfer date above.
    If glbCompSerial = "S/N - 2205W" Then 'Crown Investment Corp. - Ticket #14090
        If rsBF("BF_BCODE") = "GLFE" Then
            SQLQ = "UPDATE HRBENFT SET BF_AMT=(" & xSalary & " * BF_FACTOR) - 25000 WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        Else
            SQLQ = "UPDATE HRBENFT SET BF_AMT=" & xSalary & " * BF_FACTOR WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        End If
    Else
        SQLQ = "UPDATE HRBENFT SET BF_AMT=" & xSalary & " * BF_FACTOR WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
    End If
    gdbAdoIhr001.Execute SQLQ
    
    If xSalary <> 0 Then
        SQLQ = "UPDATE HRBENFT SET BF_ROUND = 0 WHERE BF_ROUND IS NULL AND " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
    
        SQLQ = "UPDATE HRBENFT SET BF_AMT=BF_MINIMUM WHERE BF_MINIMUM<>0 AND BF_MINIMUM>BF_AMT AND " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "UPDATE HRBENFT SET BF_AMT=BF_MAXIMUM WHERE BF_MAXIMUM<>0 AND BF_MAXIMUM<BF_AMT AND " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
        
        If glbSQL Then
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=ROUND(BF_AMT/(CASE WHEN BF_ROUND=0 THEN 0.01 ELSE BF_ROUND END)+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0)* (CASE WHEN BF_ROUND=0 THEN 0.01 ELSE BF_ROUND END)"
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0) * BF_ROUND END)"
            'Hemu - Ceiling testing - Need confirmation from County of Wellington that it works fine now.
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0, (CASE WHEN BF_ROUND = 1 THEN 1 ELSE 0 END)) * BF_ROUND END)"
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT WHEN BF_ROUND=1 THEN (CEILING(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0 END)) * BF_ROUND) ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0, (CASE WHEN BF_ROUND = 1 THEN 1 ELSE 0 END)) * BF_ROUND END)"
            'Ticket #14857
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT WHEN BF_ROUND=1 THEN (CEILING(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0 END)) * BF_ROUND) ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0) * BF_ROUND END)"
            'Ticket #15270 The following sql state caused incorrect round issue
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT WHEN BF_ROUND=1 THEN (CEILING(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0 END)) * BF_ROUND) ELSE FLOOR(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END)) * BF_ROUND END)"
            
            'Ticket #18465 - change in the NEXT logic. If the Coverage Amount is evenly divisible by Rounding factor then do not round up to NEXT rounding factor.
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT WHEN BF_ROUND=1 THEN (CEILING(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0 END)) * BF_ROUND) ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0) * BF_ROUND END)
            'Ticket #20315
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT WHEN BF_ROUND=1 THEN (CEILING(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0 END)) * BF_ROUND) ELSE (CASE WHEN BF_NEXTNEAREST = 'N' AND ((BF_AMT/BF_ROUND) - CAST(FLOOR(BF_AMT/BF_ROUND) AS NUMERIC)) = 0 THEN ROUND(BF_AMT/BF_ROUND,0) * BF_ROUND ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0) * BF_ROUND END) END)"
            'Ticket #21184
            'SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT WHEN BF_ROUND=1 THEN (ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0 END),0) * BF_ROUND) ELSE (CASE WHEN BF_NEXTNEAREST = 'N' AND ((BF_AMT/BF_ROUND) - CAST(FLOOR(BF_AMT/BF_ROUND) AS NUMERIC)) = 0 THEN ROUND(BF_AMT/BF_ROUND,0) * BF_ROUND ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0) * BF_ROUND END) END)"
            SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT WHEN BF_ROUND=1 AND BF_NEXTNEAREST = 'R' THEN (ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0 END),0) * BF_ROUND) ELSE (CASE WHEN BF_NEXTNEAREST = 'N' AND ((BF_AMT/BF_ROUND) - CAST(FLOOR(BF_AMT/BF_ROUND) AS NUMERIC)) = 0 THEN ROUND(BF_AMT/BF_ROUND,0) * BF_ROUND ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0) * BF_ROUND END) END)"
            
        ElseIf glbOracle Then
            SQLQ = "UPDATE HRBENFT SET BF_AMT=(CASE WHEN BF_ROUND=0 THEN BF_AMT ELSE ROUND(BF_AMT/BF_ROUND+(CASE WHEN BF_NEXTNEAREST = 'R' THEN 0 ELSE 0.5 END),0) * BF_ROUND END)"
        Else
            SQLQ = "UPDATE HRBENFT SET BF_AMT=ROUND(BF_AMT/IIf(BF_ROUND=0,0.01,BF_ROUND)+IIf(BF_NEXTNEAREST = 'R', 0,0.5))* IIf(BF_ROUND=0,0.01,BF_ROUND)"
        End If
        SQLQ = SQLQ & " WHERE BF_AMT<>0 AND " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
        
        'Ticket #14579 - Begin The BF_AMT exceed the Maximum after the update above, so it should keep in the range
        SQLQ = "UPDATE HRBENFT SET BF_AMT=BF_MINIMUM WHERE BF_MINIMUM<>0 AND BF_MINIMUM>BF_AMT AND " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "UPDATE HRBENFT SET BF_AMT=BF_MAXIMUM WHERE BF_MAXIMUM<>0 AND BF_MAXIMUM<BF_AMT AND " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
        'Ticket #14579 - End
        
    End If
    
    If xOmersFormula And rsBF("BF_BCODE") = "OMER" Then 'Ticket #20872 Franks 09/27/2011
        'updatge total cost of OMER
        Call EmpOmersCalculate(updEMPID, rsBF("BF_BCODE"), "Y", xSalary, rsBF("BF_BENE_ID"))
    Else
        SQLQ = "UPDATE HRBENFT SET BF_TCOST ="
        If glbSQL Or glbOracle Then
            If glbCompSerial = "S/N - 2439W" And rsBF("BF_BCODE") = "STD" Then 'OK Tire Ticket #22580 Franks 09/27/2012 - Add /52
                SQLQ = SQLQ & " (CASE WHEN BF_PER IS NULL THEN BF_TCOST WHEN BF_PER =0 THEN BF_TCOST ELSE ( BF_AMT * BF_UNITCOST )/ BF_PER / 52 END) "
            Else
                SQLQ = SQLQ & " (CASE WHEN BF_PER IS NULL THEN BF_TCOST WHEN BF_PER =0 THEN BF_TCOST ELSE ( BF_AMT * BF_UNITCOST )/ BF_PER END) "
            End If
        Else
            SQLQ = SQLQ & " IIf(IsNull(BF_PER),BF_TCOST,IIf(BF_PER=0,BF_TCOST,(BF_AMT*BF_UNITCOST)/BF_PER)) "
        End If
        SQLQ = SQLQ & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
    End If
    
    If glbCompSerial = "S/N - 2262W" Then 'Wellington - Ticket #10718
        If rsBF("BF_BCODE") = "5ADB" Or rsBF("BF_BCODE") = "5GRB" Or rsBF("BF_BCODE") = "5LTB" Or _
            rsBF("BF_BCODE") = "6ADB" Or rsBF("BF_BCODE") = "6GRB" Or rsBF("BF_BCODE") = "6LTB" Or _
            rsBF("BF_BCODE") = "8GRB" Or rsBF("BF_BCODE") = "4ADW" Or rsBF("BF_BCODE") = "4GRW" Or _
            rsBF("BF_BCODE") = "4LTW" Or rsBF("BF_BCODE") = "1GRB" Then
                
            SQLQ = "UPDATE HRBENFT SET BF_TCOST ="
            If glbSQL Or glbOracle Then
                SQLQ = SQLQ & " ROUND(BF_TCOST,2)"
            Else
                SQLQ = SQLQ & " ROUND(BF_TCOST,2)"
            End If
            SQLQ = SQLQ & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    
    If CostINFO.Type = "M" Or CostINFO.Type = "W" Then     'Ticket #25235 - For weekly too * 12 even though the Covrg Amt is Weekly
        SQLQ = "UPDATE HRBENFT SET BF_TCOST =BF_TCOST *12"
        SQLQ = SQLQ & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
        gdbAdoIhr001.Execute SQLQ
    'Ticket #25235 - This is not working so Jerry and I decided to use the above * 12 which gives the right result for the client
    'ElseIf CostINFO.Type = "W" Then     'Ticket #22682 - Release 8.0 - added Weekly option to Benefit Costing
    '    SQLQ = "UPDATE HRBENFT SET BF_TCOST =BF_TCOST * 52"
    '    SQLQ = SQLQ & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
    '    gdbAdoIhr001.Execute SQLQ
    End If
    If glbCompSerial = "S/N - 2387W" Then 'Bird Packaging Limited 'Ticket #13701
        If rsBF("BF_BCODE") = "PEN" Then
            Dim rsTemSal As New ADODB.Recordset
            SQLQ = "SELECT SH_SALCD FROM HR_SALARY_HISTORY WHERE SH_CURRENT <>0 AND SH_EMPNBR = " & updEMPID & " "
            rsTemSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemSal.EOF Then
                SQLQ = "UPDATE HRBENFT SET BF_PPAMT = "
                If rsTemSal("SH_SALCD") = "A" Then
                    SQLQ = SQLQ & (Round(xSalary / 26, 2)) & " "
                Else 'Hourly
                    SQLQ = SQLQ & (Round(xSalary / 52, 2)) & " "
                End If
                SQLQ = SQLQ & " WHERE " & SQLW & " AND BF_BCODE='" & rsBF("BF_BCODE") & "'"
                gdbAdoIhr001.Execute SQLQ
            End If
            rsTemSal.Close
        End If
    End If
    
    rsBF.MoveNext
Loop
'gdbAdoIhr001.CommitTrans

If SalaryDependantFind Then
    ''Ticket #23729 Franks 05/13/2013 - comment these code, open a record to loop
    ''gdbAdoIhr001.BeginTrans
    ''SQLQ = "UPDATE HRBENFT SET "
    ''SQLQ = SQLQ & "BF_ECOST = BF_TCOST * BF_PCE, "
    ''SQLQ = SQLQ & "BF_CCOST = BF_TCOST * BF_PCC, "
    ''SQLQ = SQLQ & "BF_LUSER = '999999998'"
    ''SQLQ = SQLQ & " WHERE " & SQLW
    ''gdbAdoIhr001.Execute SQLQ
    ''
    ''If (glbCompSerial <> "S/N - 2335W") Or (glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") > 0) Then
    ''    SQLQ = "UPDATE HRBENFT SET BF_MTHCCOST ="
    ''    SQLQ = SQLQ & " ROUND(BF_CCOST  / 12, 4),BF_MTHECOST = ROUND( BF_ECOST  / 12, 4) "
    ''    SQLQ = SQLQ & " WHERE " & SQLW
    ''    gdbAdoIhr001.Execute SQLQ
    ''End If
    ''
    ''
    '''Added by Bryan Aug 21, 2007 Ticket#13546
    ''SQLQ = "UPDATE HRBENFT SET BF_TCOST = BF_ECOST + BF_CCOST WHERE BF_PCE + BF_PCC <> 1 AND " & SQLW
    ''gdbAdoIhr001.Execute SQLQ
    ''
    ''gdbAdoIhr001.CommitTrans
    
    'Ticket #23729 Franks 05/13/2013 - begin
    'comment these code, open a record to loop
    ''all benefit codes, the program needs to know if there is change on benefit,
    ''if no change the do not add record to HRAUDIT
    xIfDatChg = False
    glbBenChanged = ""  'Release 8.1
    
    SQLQ = "SELECT * FROM HRBENFT WHERE " & SQLW  '& " AND BF_LUSER = '999999998' "
    If rsBenWrk.State <> 0 Then rsBenWrk.Close
    rsBenWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsBenWrk.EOF 'xPreVal
        If IsNull(rsBenWrk("BF_ECOST")) Then xPreVal = 0 Else xPreVal = rsBenWrk("BF_ECOST")
        rsBenWrk("BF_ECOST") = rsBenWrk("BF_TCOST") * rsBenWrk("BF_PCE")
        If Not Round(xPreVal, 2) = Round(rsBenWrk("BF_ECOST"), 2) Then xIfDatChg = True
        
        If IsNull(rsBenWrk("BF_CCOST")) Then xPreVal = 0 Else xPreVal = rsBenWrk("BF_CCOST")
        rsBenWrk("BF_CCOST") = rsBenWrk("BF_TCOST") * rsBenWrk("BF_PCC")
        If Not Round(xPreVal, 2) = Round(rsBenWrk("BF_CCOST"), 2) Then xIfDatChg = True
        
        If (glbCompSerial <> "S/N - 2335W") Or (glbCompSerial = "S/N - 2335W" And InStr(1, glbSeleDiv, "HSV") > 0) Then
            If IsNull(rsBenWrk("BF_MTHCCOST")) Then xPreVal = 0 Else xPreVal = rsBenWrk("BF_MTHCCOST")
            rsBenWrk("BF_MTHCCOST") = Round((rsBenWrk("BF_CCOST") / 12), 4) 'ROUND(BF_CCOST  / 12, 4)
            If Not Round(xPreVal, 2) = Round(rsBenWrk("BF_MTHCCOST"), 2) Then xIfDatChg = True
            
            If IsNull(rsBenWrk("BF_MTHECOST")) Then xPreVal = 0 Else xPreVal = rsBenWrk("BF_MTHECOST")
            rsBenWrk("BF_MTHECOST") = Round((rsBenWrk("BF_ECOST") / 12), 4) 'ROUND(BF_CCOST  / 12, 4)
            If Not Round(xPreVal, 2) = Round(rsBenWrk("BF_MTHECOST"), 2) Then xIfDatChg = True
        End If
        
        'SQLQ = "UPDATE HRBENFT SET BF_TCOST = BF_ECOST + BF_CCOST WHERE BF_PCE + BF_PCC <> 1 AND " & SQLW
        If Not rsBenWrk("BF_PCE") + rsBenWrk("BF_PCC") = 1 Then
            rsBenWrk("BF_TCOST") = rsBenWrk("BF_ECOST") + rsBenWrk("BF_CCOST")
        End If
        rsBenWrk("BF_LUSER") = "999999998"
        rsBenWrk.Update
        
        'Release 8.1 - Save the Benefit Codes so an email notification can be sent
        glbBenChanged = glbBenChanged & rsBenWrk("BF_BCODE") & ","
        
        rsBenWrk.MoveNext
    Loop
    rsBenWrk.Close
    
    If xIfDatChg Then
        Call AUDITBENF("M")
    End If
    'Ticket #23729 Franks 05/13/2013 - end

    
    gdbAdoIhr001.BeginTrans
    SQLQ = "UPDATE HRBENFT SET "
    SQLQ = SQLQ & "BF_LUSER = '999999999'"
    SQLQ = SQLQ & " WHERE " & SQLW
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
    
    If glbWFC Then 'Ticket #15818
        Call WFCCNDBeneAuditFlag(updEMPID)
    End If
End If

rsBF.Close

'Comment by Frank 04/05/2004, ticket# 5984
'If glbCElgin Then
'    xSalary = CrtSalaryElgin(updEMPID)
'    If ElginBenefit(updEMPID, xSalary) Then Exit Sub
'End If
If glbCompSerial = "S/N - 2347W" Then
    Call updBenefitForSurreyPlace(updEMPID)
End If
End Sub

Public Sub CalcPP(Optional xCode As String, Optional xGroup As String)
    Dim rs As New ADODB.Recordset
    Dim rsIn As New ADODB.Recordset
    Dim SQLQ As String, WSQLQ As String
    Dim x As Boolean
    Dim I As Long, xTot As Long, oPayP As Double
    
    WSQLQ = ""
    If IsEmpty(xCode) = False Then
        If xCode <> "" Then
            WSQLQ = "AND BF_BCODE='" & xCode & "' and BF_GROUP='" & xGroup & "'"
        End If
    End If
    
    SQLQ = "SELECT BF_EMPNBR, BF_PPAMT, BF_MTHECOST, BF_ECOST, BF_GROUP, BF_BCODE, BF_LUSER, BF_LDATE, BF_LTIME FROM HRBENFT "
    SQLQ = SQLQ & "WHERE BF_EMPNBR=" & glbLEE_ID & " " & WSQLQ
    
    rs.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
    I = 0
    If Not rs.EOF Then
        xTot = rs.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    Do While Not rs.EOF
    'If rs.EOF = False And rs.BOF = False Then
        MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100: I = I + 1
        oPayP = rs("BF_PPAMT")
        Select Case rs("BF_GROUP")
        Case "GHON", "GHQC", "CAMPBELL", "CAMPBC", "GHQC113", "GHON113", "CAMPBC113"   'Ticket #18963, Ticket #24537 - more codes
            rs("BF_PPAMT") = rs("BF_ECOST") / 52
        Case Else
            'Frank 7/24/2008 Ticket #15270 - Begin
            'If (rs("BF_PPAMT") = 0 Or rs("BF_PPAMT") = "" Or IsNull(rs("BF_PPAMT"))) Then
            '    rs("BF_PPAMT") = rs("BF_MTHECOST") / 2
            'End If
            If IsNull(rs("BF_MTHECOST")) Then
                rs("BF_PPAMT") = 0
            Else
                rs("BF_PPAMT") = rs("BF_MTHECOST") / 2
            End If
            'Frank 7/24/2008 Ticket #15270 - End
        End Select
        rs("BF_LUSER") = glbUserID
        rs("BF_LTIME") = Time$
        rs("BF_LDATE") = Format(Now, "SHORT DATE")
        rs.Update
        If oPayP <> rs("BF_PPAMT") Then
            x = AUDITPP(rs("BF_EMPNBR"), rs("BF_BCODE"), rs("BF_PPAMT"))
        End If
        rs.MoveNext
    'End If
    Loop
    rs.Close

    MDIMain.panHelp(0).FloodType = 0
End Sub

Public Function AUDITPP(xEmpNbr, xCode, xAmount) As Boolean
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String, ACTX As String
Dim strFields As String
On Error GoTo AUDIT_ERR
AUDITPP = False
ACTX = "M"
rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNbr, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If
rsTB.Close

'rsTB.Open "SELECT * FROM HRBENFT WHERE BF_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenKeyset, adCmdText
'If rsTB.EOF Then GoTo MODNOUPD

'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv
rsTA("AU_PPAMT") = xAmount 'rsTB("BF_PPAMT") AU_BCODE
rsTA("AU_BCODE") = xCode

Dim rsEmp As New ADODB.Recordset
Dim SQLQ
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNbr 'glbLEE_ID
rsTA("AU_LDATE") = Date
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

MODNOUPD:
AUDITPP = True
Exit Function
AUDIT_ERR:

'glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
'If gintRollBack% = False Then Resume Next Else Unload Me

End Function

Public Function CrtSalaryLambton(xEMP)
    Dim SQLQ
    Dim rsSal As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    CrtSalaryLambton = 0
    SQLQ = "select SH_SALARY,SH_SALCD,SH_WHRS,JH_FTEHRS "

    If glbtermopen Then
        SQLQ = SQLQ & " from Term_SALARY_HISTORY INNER JOIN Term_JOB_HISTORY "
        SQLQ = SQLQ & " ON Term_SALARY_HISTORY.SH_JOB=Term_JOB_HISTORY.JH_JOB "
        SQLQ = SQLQ & " AND Term_SALARY_HISTORY.TERM_SEQ=Term_JOB_HISTORY.TERM_SEQ "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND JH_CURRENT<>0 AND Term_SALARY_HISTORY.TERM_SEQ=" & xEMP
        SQLQ = SQLQ & " ORDER BY JH_USRCHECK DESC "
        rsSal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Else
        SQLQ = SQLQ & " from HR_SALARY_HISTORY INNER JOIN HR_JOB_HISTORY "
        SQLQ = SQLQ & " ON HR_SALARY_HISTORY.SH_JOB=HR_JOB_HISTORY.JH_JOB "
        SQLQ = SQLQ & " AND HR_SALARY_HISTORY.SH_EMPNBR=HR_JOB_HISTORY.JH_EMPNBR "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND JH_CURRENT<>0 AND HR_SALARY_HISTORY.SH_EMPNBR=" & xEMP
        SQLQ = SQLQ & " ORDER BY JH_USRCHECK DESC "
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If

    
    If Not rsSal.EOF Then
        If Not IsNull(rsSal("SH_WHRS")) Then CrtSalaryLambton = rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52
        If Not IsNull(rsSal("JH_FTEHRS")) Then
            If rsSal("JH_FTEHRS") <> 0 Then CrtSalaryLambton = rsSal("SH_SALARY") * rsSal("JH_FTEHRS")
        End If
    End If
    rsSal.Close
End Function

Public Function CrtSalary(xEMP)
    Dim SQLQ
    Dim rsSal As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim xMultiSalSum As Double
    Dim xCrtSalary As Double
    
    If glbLambton Then
        CrtSalary = CrtSalaryLambton(xEMP)
        Exit Function
    End If
        
    xMultiSalSum = 0
    CrtSalary = 0
    xCrtSalary = 0
    glbSalaryEDate = ""
    
    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
        SQLQ = "select SH_SALARY,SH_SALCD,SH_WHRS,SH_NFAC_SALARY,SH_EDATE,SH_JOB,SH_SDATE "
    Else
        SQLQ = "select SH_SALARY,SH_SALCD,SH_WHRS,SH_EDATE,SH_JOB,SH_SDATE "
    End If

    If glbtermopen Then
        SQLQ = SQLQ & " from Term_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND TERM_SEQ=" & xEMP
        rsSal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Else
        SQLQ = SQLQ & " from HR_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & xEMP
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If
   
'if Jerry asks to only extract Multi Position client's Salary records with corresponding Acting Position = "YES"
'and sh_empnbr in (select jh_empnbr from hr_job_history
'where sh_job = jh_job and sh_sdate = jh_sdate and sh_current = jh_current and JH_POSITION_CONTROL = 'YES')
    
    If Not rsSal.EOF Then
        Do While Not rsSal.EOF
            glbSalaryEDate = rsSal("SH_EDATE")
            If rsSal("SH_SALCD") <> "H" Then
                'Frank 06/16/2004 Ticket# 6355
                If rsSal("SH_SALCD") = "A" Then
                    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
                        If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
                            xCrtSalary = rsSal("SH_NFAC_SALARY")
                        Else
                            xCrtSalary = rsSal("SH_SALARY")
                        End If
                    Else
                        xCrtSalary = rsSal("SH_SALARY")
                    End If
                ElseIf rsSal("SH_SALCD") = "M" Then ' - Monthly
                    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
                        If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
                            xCrtSalary = rsSal("SH_NFAC_SALARY") * 12
                        Else
                            xCrtSalary = rsSal("SH_SALARY") * 12
                        End If
                    Else
                        xCrtSalary = rsSal("SH_SALARY") * 12
                    End If
                'added by Bryan 26/09/05 Ticket#9354
                ElseIf rsSal("SH_SALCD") = "D" Then
                    If GetLeapYear(Year(Date)) Then
                        If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
                            If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
                                xCrtSalary = rsSal("SH_NFAC_SALARY") * 366
                            Else
                                xCrtSalary = rsSal("SH_SALARY") * 366
                            End If
                        Else
                            xCrtSalary = rsSal("SH_SALARY") * 366
                            
                            'Ticket #17654 - formula correction
                            xCrtSalary = (rsSal("SH_SALARY") / GetJHData(xEMP, "JH_DHRS", 1)) * (rsSal("SH_WHRS") * 52)
                        End If
                    Else
                        If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
                            If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
                                xCrtSalary = rsSal("SH_NFAC_SALARY") * 365
                            Else
                                xCrtSalary = rsSal("SH_SALARY") * 365
                            End If
                        Else
                            xCrtSalary = rsSal("SH_SALARY") * 365
                        
                            'Ticket #17654 - formula correction
                            xCrtSalary = (rsSal("SH_SALARY") / GetJHData(xEMP, "JH_DHRS", 1)) * (rsSal("SH_WHRS") * 52)
                        End If
                    End If
                End If
            Else
                SQLQ = "select JH_FTEHRS "
                If glbtermopen Then
                    SQLQ = SQLQ & " from Term_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND TERM_SEQ=" & xEMP
                    'Ticket #21493 - If it's a multi position then more than one Job record is retrieved
                    SQLQ = SQLQ & " AND JH_JOB = '" & rsSal("SH_JOB") & "'"
                    'Ticket #23906 Franks 06/21/2013 - don't use position start date
                    'SQLQ = SQLQ & " AND JH_SDATE = " & Date_SQL(rsSal("SH_SDATE"))
                    SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
                    rsJOB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset
                Else
                    SQLQ = SQLQ & " from HR_JOB_HISTORY "
                    SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEMP
                    'Ticket #21493 - If it's a multi position then more than one Job record is retrieved
                    SQLQ = SQLQ & " AND JH_JOB = '" & rsSal("SH_JOB") & "'"
                    'Ticket #23906 Franks 06/21/2013 - don't use position start date
                    'SQLQ = SQLQ & " AND JH_SDATE = " & Date_SQL(rsSal("SH_SDATE"))
                    SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
                    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
                End If
                If Not rsJOB.EOF Then
                    If Not IsNull(rsSal("SH_WHRS")) Then
                        If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
                            If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
                                xCrtSalary = rsSal("SH_NFAC_SALARY") * rsSal("SH_WHRS") * 52
                            Else
                                xCrtSalary = rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52
                            End If
                        Else
                            xCrtSalary = rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52
                        End If
                    End If
                    
                    If Not IsNull(rsJOB("JH_FTEHRS")) Then
                        If rsJOB("JH_FTEHRS") <> 0 Then
                            If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
                                If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
                                    xCrtSalary = rsSal("SH_NFAC_SALARY") * rsJOB("JH_FTEHRS")
                                Else
                                    xCrtSalary = rsSal("SH_SALARY") * rsJOB("JH_FTEHRS")
                                End If
                            Else
                                xCrtSalary = rsSal("SH_SALARY") * rsJOB("JH_FTEHRS")
                            End If
                        End If
                    End If
                End If
                rsJOB.Close
            End If
            
            'Ticket #20347 - Sum up all the current salaries
            xMultiSalSum = xMultiSalSum + xCrtSalary
            rsSal.MoveNext
        Loop
    End If
    rsSal.Close
    
    CrtSalary = xMultiSalSum
    
End Function

'Public Function CrtSalary(xEMP)
'    Dim SQLQ
'    Dim rsSal As New ADODB.Recordset
'    Dim rsJOB As New ADODB.Recordset
'
'    If glbLambton Then
'        CrtSalary = CrtSalaryLambton(xEMP)
'        Exit Function
'    End If
'
'    xMultiSalSum = 0
'    CrtSalary = 0
'    glbSalaryEDate = ""
'
'    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
'        SQLQ = "select SH_SALARY,SH_SALCD,SH_WHRS,SH_NFAC_SALARY,SH_EDATE "
'    Else
'        SQLQ = "select SH_SALARY,SH_SALCD,SH_WHRS,SH_EDATE "
'    End If
'
'    If glbtermopen Then
'        SQLQ = SQLQ & " from Term_SALARY_HISTORY "
'        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND TERM_SEQ=" & xEMP
'        rsSal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
'    Else
'        SQLQ = SQLQ & " from HR_SALARY_HISTORY "
'        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & xEMP
'        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
'    End If
'
'    If Not rsSal.EOF Then
'        glbSalaryEDate = rsSal("SH_EDATE")
'        If rsSal("SH_SALCD") <> "H" Then
'            'Frank 06/16/2004 Ticket# 6355
'            If rsSal("SH_SALCD") = "A" Then
'                If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
'                    If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
'                        CrtSalary = rsSal("SH_NFAC_SALARY")
'                    Else
'                        CrtSalary = rsSal("SH_SALARY")
'                    End If
'                Else
'                    CrtSalary = rsSal("SH_SALARY")
'                End If
'            ElseIf rsSal("SH_SALCD") = "M" Then ' - Monthly
'                If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
'                    If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
'                        CrtSalary = rsSal("SH_NFAC_SALARY") * 12
'                    Else
'                        CrtSalary = rsSal("SH_SALARY") * 12
'                    End If
'                Else
'                    CrtSalary = rsSal("SH_SALARY") * 12
'                End If
'            'added by Bryan 26/09/05 Ticket#9354
'            ElseIf rsSal("SH_SALCD") = "D" Then
'                If GetLeapYear(Year(Date)) Then
'                    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
'                        If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
'                            CrtSalary = rsSal("SH_NFAC_SALARY") * 366
'                        Else
'                            CrtSalary = rsSal("SH_SALARY") * 366
'                        End If
'                    Else
'                        CrtSalary = rsSal("SH_SALARY") * 366
'
'                        'Ticket #17654 - formula correction
'                        CrtSalary = (rsSal("SH_SALARY") / GetJHData(xEMP, "JH_DHRS", 1)) * (rsSal("SH_WHRS") * 52)
'                    End If
'                Else
'                    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
'                        If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
'                            CrtSalary = rsSal("SH_NFAC_SALARY") * 365
'                        Else
'                            CrtSalary = rsSal("SH_SALARY") * 365
'                        End If
'                    Else
'                        CrtSalary = rsSal("SH_SALARY") * 365
'
'                        'Ticket #17654 - formula correction
'                        CrtSalary = (rsSal("SH_SALARY") / GetJHData(xEMP, "JH_DHRS", 1)) * (rsSal("SH_WHRS") * 52)
'                    End If
'                End If
'            End If
'        Else
'            SQLQ = "select JH_FTEHRS "
'            If glbtermopen Then
'                SQLQ = SQLQ & " from Term_JOB_HISTORY "
'                SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND TERM_SEQ=" & xEMP
'                rsJOB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset
'            Else
'                SQLQ = SQLQ & " from HR_JOB_HISTORY "
'                SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEMP
'                rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
'            End If
'            If Not rsJOB.EOF Then
'                If Not IsNull(rsSal("SH_WHRS")) Then
'                    If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
'                        If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
'                            CrtSalary = rsSal("SH_NFAC_SALARY") * rsSal("SH_WHRS") * 52
'                        Else
'                            CrtSalary = rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52
'                        End If
'                    Else
'                        CrtSalary = rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52
'                    End If
'                End If
'
'                If Not IsNull(rsJOB("JH_FTEHRS")) Then
'                    If rsJOB("JH_FTEHRS") <> 0 Then
'                        If glbCompSerial = "S/N - 2363W" Then   'City of Kawartha Lakes - Ticket #13179
'                            If rsSal("SH_NFAC_SALARY") <> 0 And Not IsNull(rsSal("SH_NFAC_SALARY")) Then
'                                CrtSalary = rsSal("SH_NFAC_SALARY") * rsJOB("JH_FTEHRS")
'                            Else
'                                CrtSalary = rsSal("SH_SALARY") * rsJOB("JH_FTEHRS")
'                            End If
'                        Else
'                            CrtSalary = rsSal("SH_SALARY") * rsJOB("JH_FTEHRS")
'                        End If
'                    End If
'                End If
'            End If
'            rsJOB.Close
'        End If
'    End If
'    rsSal.Close
'
'End Function

Public Function CrtBeneCost(updEMPID, xSalary, xGroup, xBCode) As BenefitCost
    Dim tSalary
    Dim SQLQ
    Dim rsCost As New ADODB.Recordset
    Dim PreviousMax
    Dim FTE
    
    On Error Resume Next
    
    CrtBeneCost.Salary = xSalary
    CrtBeneCost.Type = "A"
    
    SQLQ = "SELECT * FROM HR_BENEFIT_COST "
    SQLQ = SQLQ & " WHERE CU_BCODE='" & xBCode & "'"
    SQLQ = SQLQ & " AND (CU_BENEFIT_GROUP='" & xGroup & "' OR CU_BENEFIT_GROUP='' OR CU_BENEFIT_GROUP IS NULL)"
    rsCost.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsCost.EOF Then
        CrtBeneCost.Salary = xSalary
        CrtBeneCost.Type = "A"
    Else
        tSalary = 0
        PreviousMax = 0
        'Ticket #22682 - Release 8.0 - added Weekly option to Benefit Costing
        xSalary = xSalary / IIf(rsCost("CU_TYPE") = "M", 12, IIf(rsCost("CU_TYPE") = "W", 52, 1))
'        FTE = GetJHData(updEMPID, "JH_FTENUM", 1)
'        xSalary = IIf(rsCost("CU_FTE") <> 0, xSalary * FTE, xSalary)
        CrtBeneCost.Type = rsCost("CU_TYPE")
        Do Until rsCost.EOF
            'Ticket #27448 - Let's Talk Science
            If glbCompSerial = "S/N - 2353W" Then
                tSalary = tSalary + rsCost("CU_PCT") * (xSalary - PreviousMax)
                If tSalary >= rsCost("CU_MAX") Then
                    tSalary = rsCost("CU_MAX")
                End If
            Else
                If xSalary >= rsCost("CU_MAX") Then
                    tSalary = tSalary + rsCost("CU_PCT") * (rsCost("CU_MAX") - PreviousMax)
                ElseIf xSalary > PreviousMax And xSalary <= rsCost("CU_MAX") Then
                    tSalary = tSalary + rsCost("CU_PCT") * (xSalary - PreviousMax)
                End If
            End If
            PreviousMax = rsCost("CU_MAX")
            rsCost.MoveNext
        Loop
        CrtBeneCost.Salary = tSalary
        
    End If
End Function


Public Function CrtSalaryElgin(xEMP)
    Dim SQLQ
    Dim rsSal As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    CrtSalaryElgin = 0
    SQLQ = "select SH_SALARY,SH_SALCD,SH_WHRS "

    Dim xDate, xYY
    xYY = Year(Now) - 1
    xDate = CVDate(GetMonth("Dec") & " 1," & Str(xYY))
    If glbtermopen Then
        SQLQ = SQLQ & " from Term_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & xEMP
        SQLQ = SQLQ & " And SH_EDATE <= " & Date_SQL(xDate)
        SQLQ = SQLQ & "ORDER BY SH_EDATE DESC "
        rsSal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Else
        SQLQ = SQLQ & " from HR_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_EMPNBR=" & xEMP
        SQLQ = SQLQ & " And SH_EDATE <= " & Date_SQL(xDate)
        SQLQ = SQLQ & "ORDER BY SH_EDATE DESC "
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    End If

    
    
    If Not rsSal.EOF Then
        If rsSal("SH_SALCD") <> "H" Then
            'CrtSalaryElgin = rsSAL("SH_SALARY")
            'Frank 06/16/2004 Ticket# 6355
            If rsSal("SH_SALCD") = "A" Then
                CrtSalaryElgin = rsSal("SH_SALARY")
            ElseIf rsSal("SH_SALCD") = "M" Then ' - Monthly
                CrtSalaryElgin = rsSal("SH_SALARY") * 12
            'added by Bryan 26/09/05 Ticket#9343
            ElseIf rsSal("SH_SALCD") = "D" Then
                If GetLeapYear(Year(Date)) Then
                    CrtSalaryElgin = rsSal("SH_SALARY") * 366
                Else
                    CrtSalaryElgin = rsSal("SH_SALARY") * 365
                End If
            End If
        Else
            SQLQ = "select JH_FTEHRS "
            If glbtermopen Then
                SQLQ = SQLQ & " from Term_JOB_HISTORY "
                SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND TERM_SEQ=" & xEMP
                rsJOB.Open SQLQ, gdbAdoIhr001X, adOpenKeyset
            Else
                SQLQ = SQLQ & " from HR_JOB_HISTORY "
                SQLQ = SQLQ & " WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xEMP
                rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset
            End If
            If Not rsJOB.EOF Then
                If Not IsNull(rsSal("SH_WHRS")) Then CrtSalaryElgin = rsSal("SH_SALARY") * rsSal("SH_WHRS") * 52
                If Not IsNull(rsJOB("JH_FTEHRS")) Then
                    If rsJOB("JH_FTEHRS") <> 0 Then CrtSalaryElgin = rsSal("SH_SALARY") * rsJOB("JH_FTEHRS")
                End If
            End If
            rsJOB.Close
        End If
    End If
    rsSal.Close
End Function

Public Function Get_Fields(db As ADODB.Connection, TableName As String, UnField As String)
Dim rsTB As New ADODB.Recordset
Dim FdList As String
Dim x As Integer

Set rsTB = Nothing
'Hemu - Ticket #16535 - Trying to optimize the process. The line below is retrieving records in the table.
'So e.g. HR_ATTENDANCE normally large table - will have all it's records retrieved. We do not need that since
'this function requires column names only
'rsTB.Open TableName, db, adOpenForwardOnly
rsTB.Open "SELECT * FROM " & TableName & " WHERE 1=2", db, adOpenForwardOnly
FdList = ""
For x = 0 To rsTB.Fields.count - 1
    If UnField = "" Or InStr(UCase(UnField) & ",", UCase(rsTB.Fields(x).name) & ",") = 0 Then
        FdList = FdList & ", " & rsTB.Fields(x).name
    End If
Next
Get_Fields = Mid(FdList, 2)
rsTB.Close
Set rsTB = Nothing
End Function

Public Function GetCode_Data(xtbl, xKey, xField, Optional xDefault)
    Dim rsTable As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT " & xField & " FROM HRTABL WHERE TB_NAME = '" & xtbl & "' AND TB_KEY = '" & xKey & "'"
    rsTable.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTable.EOF Then
        GetCode_Data = rsTable(xField)
    Else
        If Not IsMissing(xDefault) Then
            GetCode_Data = xDefault
        Else
            GetCode_Data = ""
        End If
    End If
    rsTable.Close
    Set rsTable = Nothing
    
End Function

Public Sub glbEmpWrk(xEmpList, xDate1, xDate2)
Dim SQLX
Dim xRecAffected As Long
Dim rsTTemp As New ADODB.Recordset
Dim SQLQ As String, xEmpNbr

If glbOracle Then
    Call glbEmpWrkOracle(xEmpList, xDate1, xDate2)
    Exit Sub
End If

If Not glbSQL Then
    Call glbEmpWrkAccess(xEmpList, xDate1, xDate2)
    Exit Sub
End If

MDIMain.panHelp(0).FloodPercent = 10
    '--------------------------------------------------- 01OK
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_NUMERIC,TT_FTE,TT_FTEHRS,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT ED_COMPNO AS TT_COMPNO, ED_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'01' AS TT_RECNBR,"
  
    
    SQLX = SQLX & "DateDiff(day,ED_DOH, getdate()) / 365 AS TT_NUMERIC,"
    SQLX = SQLX & "DateDiff(day,ED_DOB, getdate()) / 365 AS TT_FTE, "
    SQLX = SQLX & " (case when ED_SENDTE is NULL then NULL else DateDiff(day,ED_SENDTE, getdate()) / 365 end) AS TT_FTEHRS "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM HREMP WHERE ED_EMPNBR in " & xEmpList
    gdbAdoIhr001.Execute SQLX
'     '--------------------------------------------------- 10OK
'MDIMain.panHelp(0).FloodPercent = 15
'    SQLX = "INSERT INTO HREMPWRK "
'    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
'    SQLX = SQLX & "TT_LANG1,TT_LANG2,TT_WRKEMP) "
'    SQLX = SQLX & in_SQL(glbIHRDBW)
'    SQLX = SQLX & "SELECT ED_COMPNO AS TT_COMPNO, ED_EMPNBR AS TT_EMPNBR,"
'    SQLX = SQLX & "'10' AS TT_RECNBR,"
'    SQLX = SQLX & "T1.TB_DESC AS TT_LANG1,"
'    SQLX = SQLX & "T2.TB_DESC AS TT_LANG2 "
'    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
'    SQLX = SQLX & " FROM (HREMP as Z LEFT  JOIN HRTABL  as T1 ON (Z.ED_LANG1 = T1.TB_KEY) AND (Z.ED_LANG1_TABL =T1.TB_NAME))"
'    SQLX = SQLX & "LEFT JOIN HRTABL AS T2 ON (Z.ED_LANG2 = T2.TB_KEY) AND (Z.ED_LANG2_TABL = T2.TB_NAME)"
'    SQLX = SQLX & "WHERE ED_EMPNBR in " & xEmplist
'    gdbAdoIhr001.Execute SQLX
'George Modified on Mar 21,2006 #10574
     '--------------------------------------------------- 10OK
MDIMain.panHelp(0).FloodPercent = 15
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_LANG1,TT_LANG2,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT EL_COMPNO AS TT_COMPNO, EL_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'10' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_LANG1,"
    SQLX = SQLX & "T2.TB_DESC AS TT_LANG2 "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HR_LANGUAGE as Z LEFT  JOIN HRTABL  as T1 ON (Z.EL_LANG_SPOKEN = T1.TB_KEY) AND (Z.EL_LANG_SPOKEN_TABL =T1.TB_NAME))"
    SQLX = SQLX & "LEFT JOIN HRTABL AS T2 ON (Z.EL_LANG_WRITTEN = T2.TB_KEY) AND (Z.EL_LANG_WRITTEN_TABL = T2.TB_NAME)"
    SQLX = SQLX & "WHERE EL_EMPNBR in " & xEmpList
    gdbAdoIhr001.Execute SQLX
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '10' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_LANG1,TT_LANG2,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'10' AS TT_RECNBR,"
        SQLX = SQLX & "'No Language' AS TT_LANG1,"
        SQLX = SQLX & "Null AS TT_LANG2 "
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
'George Modified on Mar 21,2006 #10574
    '--------------------------------------------------- 03OK
MDIMain.panHelp(0).FloodPercent = 20
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_SEX,TT_NAMEFLD,TT_CHAR10,TT_DATEFLD, "
    SQLX = SQLX & "TT_OLDDEPT,TT_NEWDEPT,TT_OLDDIV, " 'Ticket #16395 for WFC COB fields
    SQLX = SQLX & "TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT DP_COMPNO AS TT_COMPNO, DP_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'03' AS TT_RECNBR,"
    SQLX = SQLX & "DP_SEX AS TT_SEX,"
    SQLX = SQLX & "DP_SNAME+','+DP_FNAME  AS TT_NAMEFLD,"
    SQLX = SQLX & "DP_RELATE AS TT_CHAR10,"
    SQLX = SQLX & "DP_DOB AS TT_DATEFLD,"
    'Ticket #16395 for WFC COB fields - begin
    SQLX = SQLX & "DP_MEDICAL AS TT_OLDDEPT,"
    SQLX = SQLX & "DP_DENTAL AS TT_NEWDEPT,"
    SQLX = SQLX & "DP_OTHER AS TT_OLDDIV"
    'Ticket #16395 for WFC COB fields - end
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM HRDEPEND WHERE DP_EMPNBR IN " & xEmpList
    gdbAdoIhr001.Execute SQLX
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '03' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_SEX,TT_NAMEFLD,TT_CHAR10,TT_DATEFLD,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'03' AS TT_RECNBR,"
        SQLX = SQLX & "Null AS TT_SEX,"
        SQLX = SQLX & "'No Dependent'  AS TT_NAMEFLD,"
        SQLX = SQLX & "Null AS TT_CHAR10,"
        SQLX = SQLX & "Null AS TT_DATEFLD"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
MDIMain.panHelp(0).FloodPercent = 25
    '--------------------------------------------------- 05, 08
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_COMPA,TT_JOB,TT_GRID,TT_SALARY,TT_SALCD,TT_SEDATE,TT_SALR1,TT_SALPC1,TT_NEXTDAT,TT_WHRS,TT_GRADE,TT_DATECHG,TT_SALR2,TT_SALR3,TT_SALPC2,TT_SALPC3,TT_JOBCODE,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT SH_COMPNO AS TT_COMPNO, SH_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "(CASE WHEN SH_CURRENT<>0 THEN '05' ELSE '08' END) AS TT_RECNBR,"
    SQLX = SQLX & "(CASE WHEN SH_CURRENT<>0 THEN NULL ELSE SH_COMPA END) AS TT_COMPA,"
    SQLX = SQLX & "J.JB_DESCR AS TT_JOB,"
    SQLX = SQLX & "T4.TB_DESC AS TT_GRID,"
    SQLX = SQLX & "SH_SALARY AS TT_SALARY,"
    SQLX = SQLX & "SH_SALCD AS TT_SALCD,"
    SQLX = SQLX & "SH_EDATE AS TT_SEDATE,"
    SQLX = SQLX & "T1.TB_DESC AS TT_SALR1,"
    SQLX = SQLX & "SH_SALPC1 AS TT_SALPC1,"
    SQLX = SQLX & "SH_NEXTDAT AS TT_NEXTDAT,"
    SQLX = SQLX & "SH_WHRS AS TT_WHRS,"
    SQLX = SQLX & "SH_GRADE AS TT_GRADE,"
    SQLX = SQLX & "SH_SDATE AS TT_DATECHG,"
    SQLX = SQLX & "T2.TB_DESC AS TT_SALR2,"
    SQLX = SQLX & "T3.TB_DESC AS TT_SALR3,"
    SQLX = SQLX & "SH_SALPC2 AS TT_SALPC2,"
    SQLX = SQLX & "SH_SALPC3 AS TT_SALPC3, "
    SQLX = SQLX & "SH_JOB AS TT_JOBCODE "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM ((((HR_SALARY_HISTORY as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.SH_SREAS1 = T1.TB_KEY) AND (Z.SH_SREAS_TABLE = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.SH_SREAS2 = T2.TB_KEY) AND (Z.SH_SREAS_TABLE = T2.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T3 ON (Z.SH_SREAS3 = T3.TB_KEY) AND (Z.SH_SREAS_TABLE = T3.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T4 ON (Z.SH_GRID = T4.TB_KEY) AND (Z.SH_GRID_TABL = T4.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRJOB AS J ON (Z.SH_JOB = J.JB_CODE) "
    SQLX = SQLX & " WHERE SH_EMPNBR IN " & xEmpList
    
    'Hemu
    SQLX = SQLX & " ORDER BY SH_EDATE DESC "
    'Hemu
    'added by Bryan 22/Sep/05 Ticket#9343
    If glbCompSerial = "S/N - 2373W" Then 'Muskoka
        SQLX = Replace(SQLX, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
    End If
gdbAdoIhr001.Execute SQLX
MDIMain.panHelp(0).FloodPercent = 30
    '--------------------------------------------------- 04, 07
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_JOB,TT_GRID,TT_JREAS,TT_REPTAU,TT_DATECHG,TT_SDATE,TT_SHIFT,TT_FTE,TT_FTEHRS,TT_WHRS,TT_DHRS,TT_PHRS,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT JH_COMPNO AS TT_COMPNO, JH_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "(CASE WHEN JH_CURRENT<>0 THEN '04' ELSE '07' END) AS TT_RECNBR,"
    SQLX = SQLX & "J.JB_DESCR AS TT_JOB,"
    SQLX = SQLX & "T2.TB_DESC AS TT_GRID,"
    SQLX = SQLX & "T1.TB_DESC AS TT_JREAS,"
    SQLX = SQLX & "JH_REPTAU AS TT_REPTAU,"
    SQLX = SQLX & "JH_SDATE AS TT_DATECHG,"
    SQLX = SQLX & "JH_SDATE AS TT_SDATE,"
    SQLX = SQLX & "JH_SHIFT AS TT_SHIFT,"
    SQLX = SQLX & "JH_FTENUM AS TT_FTE,"
    SQLX = SQLX & "JH_FTEHRS AS TT_FTEHRS,"
    SQLX = SQLX & "JH_WHRS AS TT_WHRS,"
    SQLX = SQLX & "JH_DHRS AS TT_DHRS,"
    SQLX = SQLX & "JH_PHRS AS TT_PHRS "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (((HR_JOB_HISTORY as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.JH_JREASON = T1.TB_KEY) AND (Z.JH_ENDREAS_TABL = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.JH_GRID= T2.TB_KEY) AND (Z.JH_GRID_TABL = T2.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRJOB AS J ON (Z.JH_JOB = J.JB_CODE)) "
    SQLX = SQLX & " WHERE JH_EMPNBR IN " & xEmpList
    
    'Hemu
    SQLX = SQLX & " ORDER BY JH_SDATE DESC"
    'Hemu
    
    gdbAdoIhr001.Execute SQLX
    'Debug.Print SQLX
MDIMain.panHelp(0).FloodPercent = 35
    '--------------------------------------------------- 06, 09
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_JOB,TT_REPTAU,TT_PCODE,TT_PREVIEW,TT_PNEXT,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT PH_COMPNO AS TT_COMPNO, PH_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "(CASE WHEN PH_CURRENT<>0 THEN '06' ELSE '09' END) AS TT_RECNBR,"
    SQLX = SQLX & "J.JB_DESCR AS TT_JOB,"
    SQLX = SQLX & "PH_REPTAU AS TT_REPTAU,"
    SQLX = SQLX & "T1.TB_DESC AS TT_PCODE,"
    SQLX = SQLX & "PH_PREVIEW AS TT_PREVIEW,"
    SQLX = SQLX & "PH_PNEXT AS TT_PNEXT "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HR_PERFORM_HISTORY as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.PH_PCODE = T1.TB_KEY) AND (Z.PH_PCODE_TABLE = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRJOB AS J ON (Z.PH_JOB = J.JB_CODE) "
    SQLX = SQLX & " WHERE PH_EMPNBR IN " & xEmpList
    
    SQLX = SQLX & " ORDER BY PH_PREVIEW DESC, PH_PNEXT DESC"
    
    gdbAdoIhr001.Execute SQLX
MDIMain.panHelp(0).FloodPercent = 40
    '--------------------------------------------------- 12OK
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_ETYPE,TT_EFDATE,TT_ETDATE,TT_EACTUAL,TT_COEFLAG,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'12' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_ETYPE,"
    SQLX = SQLX & "FDATE AS TT_EFDATE,"
    SQLX = SQLX & "TDATE AS TT_ETDATE,"
    SQLX = SQLX & "ACT_DOLLAR AS TT_EACTUAL,"
    SQLX = SQLX & "COST_OF_EMPLOYMENT AS TT_COEFLAG "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HREARN as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.EARN_TYPE = T1.TB_KEY) AND (Z.EARN_TYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND FDATE >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND TDATE <= " & Date_SQL(xDate2)
    SQLX = SQLX & " ORDER BY TT_EFDATE DESC, TT_ETDATE DESC, EARN_TYPE "
    gdbAdoIhr001.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '12' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_ETYPE,TT_EFDATE,TT_ETDATE,TT_EACTUAL,TT_COEFLAG,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'12' AS TT_RECNBR,"
        SQLX = SQLX & "'No Other Earnings' AS TT_ETYPE,"
        SQLX = SQLX & "Null AS TT_EFDATE,"
        SQLX = SQLX & "Null AS TT_ETDATE,"
        SQLX = SQLX & "Null AS TT_EACTUAL,"
        SQLX = SQLX & "0 AS TT_COEFLAG "
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 45
  '--------------------------------------------------- 11OK
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_SKILLD,TT_EXPFACT,TT_SKLDTE,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, SE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'11' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_SKILLD,"
    SQLX = SQLX & "SE_LEVEL AS TT_EXPFACT,"
    SQLX = SQLX & "SE_DATE AS TT_SKLDTE "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HREMPSKL as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.SE_SKILL = T1.TB_KEY) AND (Z.SE_SKILL_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE SE_EMPNBR IN " & xEmpList
    
    'Hemu
    SQLX = SQLX & " ORDER BY SE_DATE DESC "
    'Hemu
    
    gdbAdoIhr001.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '11' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_SKILLD,TT_EXPFACT,TT_SKLDTE,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'11' AS TT_RECNBR,"
        SQLX = SQLX & "'No Skill' AS TT_SKILLD,"
        SQLX = SQLX & "Null AS TT_EXPFACT,"
        SQLX = SQLX & "Null AS TT_SKLDTE "
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
MDIMain.panHelp(0).FloodPercent = 50
    '--------------------------------------------------- 13OK
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_DEGREED,TT_YEAR,TT_MAJORD,TT_MINORD,TT_COMPL,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, EU_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'13' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_DEGREED,"
    SQLX = SQLX & "EU_YEAR AS TT_YEAR,"
    SQLX = SQLX & "T2.TB_DESC AS TT_MAJORD,"
    SQLX = SQLX & "T3.TB_DESC AS TT_MINORD, "
    SQLX = SQLX & "EU_COMP AS TT_COMPL"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (((HREDU as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.EU_DEGREE = T1.TB_KEY) AND (Z.EU_DEGREE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.EU_MAJOR = T2.TB_KEY) AND (Z.EU_MAJOR_TABL = T2.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T3 ON (Z.EU_MINOR = T3.TB_KEY) AND (Z.EU_MINOR_TABL = T3.TB_NAME))"
    SQLX = SQLX & " WHERE EU_EMPNBR IN " & xEmpList
    
    SQLX = SQLX & " ORDER BY EU_YEAR DESC"
    
    gdbAdoIhr001.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '13' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_DEGREED,TT_YEAR,TT_MAJORD,TT_MINORD,TT_COMPL,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'13' AS TT_RECNBR,"
        SQLX = SQLX & "'No Formal Education' AS TT_DEGREED,"
        SQLX = SQLX & "Null AS TT_YEAR,"
        SQLX = SQLX & "Null AS TT_MAJORD,"
        SQLX = SQLX & "Null AS TT_MINORD, "
        SQLX = SQLX & "Null AS TT_COMPL"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 55
  '--------------------------------------------------- 15OK
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_CTYPED,TT_COURSE,TT_DATCOMP,TT_RESULTD,TT_TBCO,TT_TBEMP,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, ES_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'15' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_CTYPED,"
    SQLX = SQLX & "ES_COURSE AS TT_COURSE,"
    SQLX = SQLX & "ES_DATCOMP AS TT_DATCOMP,"
    SQLX = SQLX & "T2.TB_DESC AS TT_RESULTD,"
    SQLX = SQLX & "ES_TBCO AS TT_TBCO,"
    SQLX = SQLX & "ES_TBEMP AS TT_TBEMP"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HREDSEM as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.ES_CTYPE = T1.TB_KEY) AND (Z.ES_CTYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.ES_RESULTS = T2.TB_KEY) AND (Z.ES_RESULTS_TABL = T2.TB_NAME)"
    SQLX = SQLX & " WHERE ES_EMPNBR IN " & xEmpList
    SQLX = SQLX & " ORDER BY ES_DATCOMP DESC, ES_CTYPE ASC"
    gdbAdoIhr001.Execute SQLX
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '15' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_CTYPED,TT_COURSE,TT_DATCOMP,TT_RESULTD,TT_TBCO,TT_TBEMP,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'15' AS TT_RECNBR,"
        SQLX = SQLX & "'No Course/Seminar' AS TT_CTYPED,"
        SQLX = SQLX & "Null AS TT_COURSE,"
        SQLX = SQLX & "Null AS TT_DATCOMP,"
        SQLX = SQLX & "Null AS TT_RESULTD,"
        SQLX = SQLX & "Null AS TT_TBCO,"
        SQLX = SQLX & "Null AS TT_TBEMP"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
MDIMain.panHelp(0).FloodPercent = 60
  '--------------------------------------------------- 17OK
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_CODED,TT_DUES,TT_COMPPD,TT_RENEWDT,TT_BEGINDT,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, TD_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'17' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_CODED,"
    SQLX = SQLX & "TD_DUES AS TT_DUES,"
    SQLX = SQLX & "TD_COMPPD AS TT_COMPPD,"
    SQLX = SQLX & "TD_RENEWDT AS TT_RENEWDT,"
    SQLX = SQLX & "TD_BEGINDT AS TT_BEGINDT"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HRTRADE as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.TD_CODE = T1.TB_KEY) AND (Z.TD_CODE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE TD_EMPNBR IN " & xEmpList
    
    SQLX = SQLX & " ORDER BY TD_BEGINDT DESC"
    
    gdbAdoIhr001.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '17' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_CODED,TT_DUES,TT_COMPPD,TT_RENEWDT,TT_BEGINDT,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'17' AS TT_RECNBR,"
        SQLX = SQLX & "'No Association/Membership' AS TT_CODED,"
        SQLX = SQLX & "Null AS TT_DUES,"
        SQLX = SQLX & "Null AS TT_COMPPD,"
        SQLX = SQLX & "Null AS TT_RENEWDT,"
        SQLX = SQLX & "Null AS TT_BEGINDT"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 65
  '--------------------------------------------------- 19OK
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_BCODED,TT_COVER,TT_BEDATE,TT_BAMT,TT_CCOST,TT_ECOST,TT_PCE,TT_PCC,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, BF_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'19' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_BCODED,"
    SQLX = SQLX & "BF_COVER AS TT_COVER,"
    SQLX = SQLX & "BF_EDATE AS TT_BEDATE,"
    SQLX = SQLX & "BF_AMT AS TT_BAMT,"
    SQLX = SQLX & "BF_CCOST AS TT_CCOST,"
    SQLX = SQLX & "BF_ECOST AS TT_ECOST,"
    SQLX = SQLX & "BF_PCE AS TT_PCE,"
    SQLX = SQLX & "BF_PCC AS TT_PCC"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HRBENFT as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.BF_BCODE = T1.TB_KEY) AND (Z.BF_BCODE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE BF_EMPNBR IN " & xEmpList
    SQLX = SQLX & " ORDER BY BF_BCODE, BF_EDATE DESC"
    gdbAdoIhr001.Execute SQLX
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '19' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_BCODED,TT_COVER,TT_BEDATE,TT_BAMT,TT_CCOST,TT_ECOST,TT_PCE,TT_PCC,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'19' AS TT_RECNBR,"
        SQLX = SQLX & "'No Benefit' AS TT_BCODED,"
        SQLX = SQLX & "Null AS TT_COVER,"
        SQLX = SQLX & "Null AS TT_BEDATE,"
        SQLX = SQLX & "Null AS TT_BAMT,"
        SQLX = SQLX & "Null AS TT_CCOST,"
        SQLX = SQLX & "Null AS TT_ECOST,"
        SQLX = SQLX & "Null AS TT_PCE,"
        SQLX = SQLX & "Null AS TT_PCC"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
MDIMain.panHelp(0).FloodPercent = 70
  '--------------------------------------------------- 20
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_BCODED,TT_BNAME,TT_BRELATE,TT_BDOB,TT_PCE,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, BD_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'20' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_BCODED,"
    SQLX = SQLX & "BD_BNAME AS TT_BNAME,"
    SQLX = SQLX & "BD_RELATE AS TT_BRELATE,"
    SQLX = SQLX & "BD_DOB AS TT_BDOB,"
    SQLX = SQLX & "BD_PC AS TT_PCE"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HRBENS as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.BD_BCODE = T1.TB_KEY) AND (Z.BD_BCODE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE BD_EMPNBR IN " & xEmpList
    SQLX = SQLX & " ORDER BY BD_BCODE"
    gdbAdoIhr001.Execute SQLX
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HREMPWRK WHERE TT_RECNBR = '20' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_BCODED,TT_BNAME,TT_BRELATE,TT_BDOB,TT_PCE,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'20' AS TT_RECNBR,"
        SQLX = SQLX & "'No Beneficiary' AS TT_BCODED,"
        SQLX = SQLX & "'' AS TT_BNAME,"
        SQLX = SQLX & "'' AS TT_BRELATE,"
        SQLX = SQLX & "Null AS TT_BDOB,"
        SQLX = SQLX & "Null AS TT_PCE"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        gdbAdoIhr001.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 75
  '--------------------------------------------------- 21
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_VAC,TT_SICK,TT_PVAC,TT_PSICK,TT_VACHRS,TT_SICKHRS,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, ED_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'21' AS TT_RECNBR,"
    SQLX = SQLX & "ED_VAC AS TT_VAC,"
    SQLX = SQLX & "ED_SICK AS TT_SICK,"
    SQLX = SQLX & "ED_PVAC AS TT_PVAC,"
    SQLX = SQLX & "ED_PSICK AS TT_PSICK,"
    SQLX = SQLX & "ED_VACT as TT_VACHRS,"
    SQLX = SQLX & "ED_SICKT as TT_SICKHRS"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM HREMP "
    SQLX = SQLX & " WHERE ED_EMPNBR IN " & xEmpList
    gdbAdoIhr001.Execute SQLX
MDIMain.panHelp(0).FloodPercent = 80
    SQLX = "UPDATE HREMPWRK SET "
    SQLX = SQLX & " HREMPWRK.TT_OTHRS =ATTHRS.OTHRS, "
    SQLX = SQLX & " HREMPWRK.TT_CTHRS = ATTHRS.CTHRS, "
    SQLX = SQLX & " HREMPWRK.TT_WCBHRS =ATTHRS.WCBHRS "
    SQLX = SQLX & " FROM HREMPWRK INNER JOIN ("
    SQLX = SQLX & " SELECT AD_EMPNBR,"
    SQLX = SQLX & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS OTHRS,"
    SQLX = SQLX & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS CTHRS,"
    SQLX = SQLX & " SUM(CASE WHEN LEFT(AD_REASON,3)='WCB' THEN AD_HRS ELSE 0 END) AS WCBHRS "
    SQLX = SQLX & " FROM HR_ATTENDANCE "
    SQLX = SQLX & " WHERE AD_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND AD_DOA >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND AD_DOA <= " & Date_SQL(xDate2)
    SQLX = SQLX & " GROUP BY AD_EMPNBR"
    SQLX = SQLX & " ) AS ATTHRS ON HREMPWRK.TT_EMPNBR=ATTHRS.AD_EMPNBR "
    SQLX = SQLX & " WHERE HREMPWRK.TT_RECNBR='21'"
    gdbAdoIhr001.Execute SQLX
   
MDIMain.panHelp(0).FloodPercent = 85
    '-----------------------------------------------------22
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_ETYPE,TT_EFDATE,TT_ETDATE,TT_ENTITLE,TT_EACTUAL,TT_COEFLAG,TT_WRKEMP) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, HE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'22' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_ETYPE,"
    SQLX = SQLX & "HE_FDATE AS TT_EFDATE,"
    SQLX = SQLX & "HE_TDATE AS TT_ETDATE,"
    SQLX = SQLX & "HE_ENTITLE AS TT_ENTITLE,"
    SQLX = SQLX & "HE_TAKEN AS TT_EACTUAL,"
    SQLX = SQLX & "HE_COE AS TT_COEFLAG"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HRENTHRS as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.HE_TYPE = T1.TB_KEY) AND (Z.HE_TYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE HE_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND HE_FDATE >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND HE_TDATE <= " & Date_SQL(xDate2)
    SQLX = SQLX & " ORDER BY HE_TYPE"
    gdbAdoIhr001.Execute SQLX
MDIMain.panHelp(0).FloodPercent = 90
  '--------------------------------------------------- 23
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_DTYPE,TT_DFDATE,TT_DTDATE,TT_DENTITL,TT_DACTUAL,TT_COEFLAG,TT_WRKEMP)"
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, DE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'23' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_DTYPE,"
    SQLX = SQLX & "DE_FDATE AS TT_DFDATE,"
    SQLX = SQLX & "DE_TDATE AS TT_DTDATE,"
    SQLX = SQLX & "DE_ENTITLE AS TT_DENTITL,"
    SQLX = SQLX & "DE_ACTUAL AS TT_DACTUAL,"
    SQLX = SQLX & "DE_COST_OF_EMPLOYMENT AS TT_COEFLAG"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    SQLX = SQLX & " FROM (HRDOLENT as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.DE_TYPE = T1.TB_KEY) AND (Z.DE_TYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE DE_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND DE_FDATE >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND DE_TDATE <= " & Date_SQL(xDate2)
    SQLX = SQLX & " ORDER BY DE_FDATE DESC"
    gdbAdoIhr001.Execute SQLX
MDIMain.panHelp(0).FloodPercent = 95
   '--------------------------------------------------- 25
    SQLX = "INSERT INTO HREMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_OLDDEPT,TT_NEWDEPT,TT_OLDDIV,TT_NEWDIV,TT_OLDEMP,TT_NEWEMP,TT_OLDPT,TT_NEWPT,TT_OLDORG,TT_NEWORG,TT_CHGDATE,TT_WRKEMP,TT_TODATE) "    'Hemu TT_TODATE
    SQLX = SQLX & in_SQL(glbIHRDBW)
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, EE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'25' AS TT_RECNBR,"
    SQLX = SQLX & "EE_OLDDEPT AS TT_OLDDEPT,"
    SQLX = SQLX & "EE_NEWDEPT AS TT_NEWDEPT,"
    SQLX = SQLX & "EE_OLDDIV AS TT_OLDDIV,"
    SQLX = SQLX & "EE_NEWDIV AS TT_NEWDIV,"
    SQLX = SQLX & "EE_OLDSTAT AS TT_OLDEMP,"
    SQLX = SQLX & "EE_NEWSTAT AS TT_NEWEMP,"
    SQLX = SQLX & "EE_OLDPT AS TT_OLDPT,"
    SQLX = SQLX & "EE_NEWPT AS TT_NEWPT,"
    SQLX = SQLX & "EE_OLDORG AS TT_OLDORG,"
    SQLX = SQLX & "EE_NEWORG AS TT_NEWORG,"
    SQLX = SQLX & "EE_CHGDATE AS TT_CHGDATE "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    
    'Hemu
    SQLX = SQLX & ",EE_TODATE AS TT_TODATE "
    'Hemu
    
    SQLX = SQLX & " FROM HREMPHIS "
    SQLX = SQLX & " WHERE EE_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND (EE_OLDDEPT IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWDEPT IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDDIV IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWDIV IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDSTAT IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWSTAT IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDORG IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWORG IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDPT IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWPT IS NOT NULL)"
    gdbAdoIhr001.Execute SQLX
  '--------------------------------------------------- 26
    If Not glbLinamar Then
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_OLDEMP,TT_NEWEMP,TT_EFDATE,TT_ETDATE,TT_JOBCODE,TT_EMP,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS SC_COMPNO, SC_EMPNBR AS TT_EMPNBR,"
        SQLX = SQLX & "'26' AS TT_RECNBR,"
        SQLX = SQLX & "SC_OLDEMP AS TT_OLDEMP,"
        SQLX = SQLX & "SC_NEWEMP AS TT_NEWEMP,"
        SQLX = SQLX & "SC_FDATE AS TT_EFDATE, "
        SQLX = SQLX & "SC_TDATE AS TT_ETDATE, "
        SQLX = SQLX & "SC_JOB AS TT_JOBCODE,"
        SQLX = SQLX & "SC_REASON AS TT_EMP,"
        SQLX = SQLX & "'" & glbUserID & "' AS TT_WRKEMP "
        SQLX = SQLX & " FROM HRSTATUS "
        SQLX = SQLX & " WHERE SC_EMPNBR IN " & xEmpList
        SQLX = SQLX & " ORDER BY SC_LDATE,SC_LTIME"
        gdbAdoIhr001.Execute SQLX
    End If
     '------------------------------------------------ end
    
    '--------------------------------------------------- 27
    'WFC Pension Beneficiary Information
    If glbWFC Then
        SQLX = "INSERT INTO HREMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_BCODED,TT_BNAME,TT_BRELATE,TT_BDOB,TT_WRKEMP) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, PE_EMPNBR AS TT_EMPNBR,"
        SQLX = SQLX & "'27' AS TT_RECNBR,"
        SQLX = SQLX & "PE_PENSIONTYPE AS TT_BCODED,"
        'SQLX = SQLX & "PE_BEN_NAME AS TT_BNAME,"
        'Ticket #24440 Franks 10/01/2013
        SQLX = SQLX & "LEFT(PE_BEN_NAME, 40) AS TT_BNAME,"
        SQLX = SQLX & "PE_BEN_RELATE AS TT_BRELATE,"
        SQLX = SQLX & "PE_BEN_DOB AS TT_BDOB"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
        SQLX = SQLX & " FROM HRP_PENSION_BENEFICIARY "
        SQLX = SQLX & " WHERE PE_EMPNBR IN " & xEmpList
        SQLX = SQLX & " ORDER BY PE_EMPNBR,PE_BEN_NAME"
        gdbAdoIhr001.Execute SQLX
    End If
    '--------------------------------------------------- end

End Sub

Public Sub glbEmpWrk_Terminate(xEmpList, xDate1, xDate2)
Dim SQLX
Dim xRecAffected As Long
Dim rsTTemp As New ADODB.Recordset
Dim SQLQ As String, xEmpNbr

'No Include Term Employees for Oracle and Access as we don't have 7.9 version for them
'If glbOracle Then
'    Call glbEmpWrkOracle(xEmplist, xDate1, xDate2)
'    Exit Sub
'End If
'
'If Not glbSQL Then
'    Call glbEmpWrkAccess(xEmplist, xDate1, xDate2)
'    Exit Sub
'End If

MDIMain.panHelp(0).FloodPercent = 10
    '--------------------------------------------------- 01OK
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_NUMERIC,TT_FTE,TT_FTEHRS,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT ED_COMPNO AS TT_COMPNO, ED_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'01' AS TT_RECNBR,"
    SQLX = SQLX & "DateDiff(day,ED_DOH, getdate()) / 365 AS TT_NUMERIC,"
    SQLX = SQLX & "DateDiff(day,ED_DOB, getdate()) / 365 AS TT_FTE, "
    SQLX = SQLX & " (case when ED_SENDTE is NULL then NULL else DateDiff(day,ED_SENDTE, getdate()) / 365 end) AS TT_FTEHRS "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, TERM_SEQ "
    SQLX = SQLX & " FROM Term_HREMP WHERE ED_EMPNBR in " & xEmpList
    gdbAdoIhr001X.Execute SQLX
'     '--------------------------------------------------- 10OK
'MDIMain.panHelp(0).FloodPercent = 15
'    SQLX = "INSERT INTO HREMPWRK "
'    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
'    SQLX = SQLX & "TT_LANG1,TT_LANG2,TT_WRKEMP) "
'    SQLX = SQLX & in_SQL(glbIHRDBW)
'    SQLX = SQLX & "SELECT ED_COMPNO AS TT_COMPNO, ED_EMPNBR AS TT_EMPNBR,"
'    SQLX = SQLX & "'10' AS TT_RECNBR,"
'    SQLX = SQLX & "T1.TB_DESC AS TT_LANG1,"
'    SQLX = SQLX & "T2.TB_DESC AS TT_LANG2 "
'    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
'    SQLX = SQLX & " FROM (HREMP as Z LEFT  JOIN HRTABL  as T1 ON (Z.ED_LANG1 = T1.TB_KEY) AND (Z.ED_LANG1_TABL =T1.TB_NAME))"
'    SQLX = SQLX & "LEFT JOIN HRTABL AS T2 ON (Z.ED_LANG2 = T2.TB_KEY) AND (Z.ED_LANG2_TABL = T2.TB_NAME)"
'    SQLX = SQLX & "WHERE ED_EMPNBR in " & xEmplist
'    gdbAdoIhr001.Execute SQLX
'George Modified on Mar 21,2006 #10574
     '--------------------------------------------------- 10OK
MDIMain.panHelp(0).FloodPercent = 15
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_LANG1,TT_LANG2,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT EL_COMPNO AS TT_COMPNO, EL_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'10' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_LANG1,"
    SQLX = SQLX & "T2.TB_DESC AS TT_LANG2 "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, TERM_SEQ "
    SQLX = SQLX & " FROM (Term_HR_LANGUAGE as Z LEFT  JOIN HRTABL  as T1 ON (Z.EL_LANG_SPOKEN = T1.TB_KEY) AND (Z.EL_LANG_SPOKEN_TABL =T1.TB_NAME)) "
    SQLX = SQLX & "LEFT JOIN HRTABL AS T2 ON (Z.EL_LANG_WRITTEN = T2.TB_KEY) AND (Z.EL_LANG_WRITTEN_TABL = T2.TB_NAME) "
    SQLX = SQLX & "WHERE EL_EMPNBR in " & xEmpList
    gdbAdoIhr001X.Execute SQLX
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR, TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '10' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_LANG1,TT_LANG2,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'10' AS TT_RECNBR,"
        SQLX = SQLX & "'No Language' AS TT_LANG1,"
        SQLX = SQLX & "Null AS TT_LANG2 "
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP," & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
'George Modified on Mar 21,2006 #10574
    '--------------------------------------------------- 03OK
MDIMain.panHelp(0).FloodPercent = 20
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_SEX,TT_NAMEFLD,TT_CHAR10,TT_DATEFLD, "
    SQLX = SQLX & "TT_OLDDEPT,TT_NEWDEPT,TT_OLDDIV, " 'Ticket #16395 for WFC COB fields
    SQLX = SQLX & "TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT DP_COMPNO AS TT_COMPNO, DP_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'03' AS TT_RECNBR,"
    SQLX = SQLX & "DP_SEX AS TT_SEX,"
    SQLX = SQLX & "DP_SNAME+','+DP_FNAME  AS TT_NAMEFLD,"
    SQLX = SQLX & "DP_RELATE AS TT_CHAR10,"
    SQLX = SQLX & "DP_DOB AS TT_DATEFLD,"
    'Ticket #16395 for WFC COB fields - begin
    SQLX = SQLX & "DP_MEDICAL AS TT_OLDDEPT,"
    SQLX = SQLX & "DP_DENTAL AS TT_NEWDEPT,"
    SQLX = SQLX & "DP_OTHER AS TT_OLDDIV"
    'Ticket #16395 for WFC COB fields - end
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, TERM_SEQ "
    SQLX = SQLX & " FROM Term_HRDEPEND WHERE DP_EMPNBR IN " & xEmpList
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR, TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '03' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_SEX,TT_NAMEFLD,TT_CHAR10,TT_DATEFLD,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'03' AS TT_RECNBR,"
        SQLX = SQLX & "Null AS TT_SEX,"
        SQLX = SQLX & "'No Dependent'  AS TT_NAMEFLD,"
        SQLX = SQLX & "Null AS TT_CHAR10,"
        SQLX = SQLX & "Null AS TT_DATEFLD"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP," & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 25
    '--------------------------------------------------- 05, 08
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_COMPA,TT_JOB,TT_GRID,TT_SALARY,TT_SALCD,TT_SEDATE,TT_SALR1,TT_SALPC1,TT_NEXTDAT,TT_WHRS,TT_GRADE,TT_DATECHG,TT_SALR2,TT_SALR3,TT_SALPC2,TT_SALPC3,TT_JOBCODE,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT SH_COMPNO AS TT_COMPNO, SH_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "(CASE WHEN SH_CURRENT<>0 THEN '05' ELSE '08' END) AS TT_RECNBR,"
    SQLX = SQLX & "(CASE WHEN SH_CURRENT<>0 THEN NULL ELSE SH_COMPA END) AS TT_COMPA,"
    SQLX = SQLX & "J.JB_DESCR AS TT_JOB,"
    SQLX = SQLX & "T4.TB_DESC AS TT_GRID,"
    SQLX = SQLX & "SH_SALARY AS TT_SALARY,"
    SQLX = SQLX & "SH_SALCD AS TT_SALCD,"
    SQLX = SQLX & "SH_EDATE AS TT_SEDATE,"
    SQLX = SQLX & "T1.TB_DESC AS TT_SALR1,"
    SQLX = SQLX & "SH_SALPC1 AS TT_SALPC1,"
    SQLX = SQLX & "SH_NEXTDAT AS TT_NEXTDAT,"
    SQLX = SQLX & "SH_WHRS AS TT_WHRS,"
    SQLX = SQLX & "SH_GRADE AS TT_GRADE,"
    SQLX = SQLX & "SH_SDATE AS TT_DATECHG,"
    SQLX = SQLX & "T2.TB_DESC AS TT_SALR2,"
    SQLX = SQLX & "T3.TB_DESC AS TT_SALR3,"
    SQLX = SQLX & "SH_SALPC2 AS TT_SALPC2,"
    SQLX = SQLX & "SH_SALPC3 AS TT_SALPC3, "
    SQLX = SQLX & "SH_JOB AS TT_JOBCODE "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM ((((Term_SALARY_HISTORY as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.SH_SREAS1 = T1.TB_KEY) AND (Z.SH_SREAS_TABLE = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.SH_SREAS2 = T2.TB_KEY) AND (Z.SH_SREAS_TABLE = T2.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T3 ON (Z.SH_SREAS3 = T3.TB_KEY) AND (Z.SH_SREAS_TABLE = T3.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T4 ON (Z.SH_GRID = T4.TB_KEY) AND (Z.SH_GRID_TABL = T4.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRJOB AS J ON (Z.SH_JOB = J.JB_CODE) "
    SQLX = SQLX & " WHERE SH_EMPNBR IN " & xEmpList
    
    'Hemu
    SQLX = SQLX & " ORDER BY SH_EDATE DESC "
    'Hemu
    'added by Bryan 22/Sep/05 Ticket#9343
    If glbCompSerial = "S/N - 2373W" Then 'Muskoka
        SQLX = Replace(SQLX, "SH_SALARY", "SH_TOTAL", 1, , vbTextCompare)
    End If
    gdbAdoIhr001X.Execute SQLX
    
MDIMain.panHelp(0).FloodPercent = 30
    '--------------------------------------------------- 04, 07
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_JOB,TT_GRID,TT_JREAS,TT_REPTAU,TT_DATECHG,TT_SDATE,TT_SHIFT,TT_FTE,TT_FTEHRS,TT_WHRS,TT_DHRS,TT_PHRS,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT JH_COMPNO AS TT_COMPNO, JH_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "(CASE WHEN JH_CURRENT<>0 THEN '04' ELSE '07' END) AS TT_RECNBR,"
    SQLX = SQLX & "J.JB_DESCR AS TT_JOB,"
    SQLX = SQLX & "T2.TB_DESC AS TT_GRID,"
    SQLX = SQLX & "T1.TB_DESC AS TT_JREAS,"
    SQLX = SQLX & "JH_REPTAU AS TT_REPTAU,"
    SQLX = SQLX & "JH_SDATE AS TT_DATECHG,"
    SQLX = SQLX & "JH_SDATE AS TT_SDATE,"
    SQLX = SQLX & "JH_SHIFT AS TT_SHIFT,"
    SQLX = SQLX & "JH_FTENUM AS TT_FTE,"
    SQLX = SQLX & "JH_FTEHRS AS TT_FTEHRS,"
    SQLX = SQLX & "JH_WHRS AS TT_WHRS,"
    SQLX = SQLX & "JH_DHRS AS TT_DHRS,"
    SQLX = SQLX & "JH_PHRS AS TT_PHRS "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (((Term_JOB_HISTORY as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.JH_JREASON = T1.TB_KEY) AND (Z.JH_ENDREAS_TABL = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.JH_GRID= T2.TB_KEY) AND (Z.JH_GRID_TABL = T2.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRJOB AS J ON (Z.JH_JOB = J.JB_CODE)) "
    SQLX = SQLX & " WHERE JH_EMPNBR IN " & xEmpList
    
    'Hemu
    SQLX = SQLX & " ORDER BY JH_SDATE DESC"
    'Hemu
    
    gdbAdoIhr001X.Execute SQLX
    'Debug.Print SQLX
    
MDIMain.panHelp(0).FloodPercent = 35
    '--------------------------------------------------- 06, 09
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_JOB,TT_REPTAU,TT_PCODE,TT_PREVIEW,TT_PNEXT,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT PH_COMPNO AS TT_COMPNO, PH_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "(CASE WHEN PH_CURRENT<>0 THEN '06' ELSE '09' END) AS TT_RECNBR,"
    SQLX = SQLX & "J.JB_DESCR AS TT_JOB,"
    SQLX = SQLX & "PH_REPTAU AS TT_REPTAU,"
    SQLX = SQLX & "T1.TB_DESC AS TT_PCODE,"
    SQLX = SQLX & "PH_PREVIEW AS TT_PREVIEW,"
    SQLX = SQLX & "PH_PNEXT AS TT_PNEXT "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_PERFORM_HISTORY as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.PH_PCODE = T1.TB_KEY) AND (Z.PH_PCODE_TABLE = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRJOB AS J ON (Z.PH_JOB = J.JB_CODE) "
    SQLX = SQLX & " WHERE PH_EMPNBR IN " & xEmpList
    
    SQLX = SQLX & " ORDER BY PH_PREVIEW DESC, PH_PNEXT DESC"
    
    gdbAdoIhr001X.Execute SQLX
    
MDIMain.panHelp(0).FloodPercent = 40
    '--------------------------------------------------- 12OK
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_ETYPE,TT_EFDATE,TT_ETDATE,TT_EACTUAL,TT_COEFLAG,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'12' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_ETYPE,"
    SQLX = SQLX & "FDATE AS TT_EFDATE,"
    SQLX = SQLX & "TDATE AS TT_ETDATE,"
    SQLX = SQLX & "ACT_DOLLAR AS TT_EACTUAL,"
    SQLX = SQLX & "COST_OF_EMPLOYMENT AS TT_COEFLAG "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_EARN as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.EARN_TYPE = T1.TB_KEY) AND (Z.EARN_TYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND FDATE >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND TDATE <= " & Date_SQL(xDate2)
    SQLX = SQLX & " ORDER BY TT_EFDATE DESC, TT_ETDATE DESC, EARN_TYPE "
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '12' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_ETYPE,TT_EFDATE,TT_ETDATE,TT_EACTUAL,TT_COEFLAG,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'12' AS TT_RECNBR,"
        SQLX = SQLX & "'No Other Earnings' AS TT_ETYPE,"
        SQLX = SQLX & "Null AS TT_EFDATE,"
        SQLX = SQLX & "Null AS TT_ETDATE,"
        SQLX = SQLX & "Null AS TT_EACTUAL,"
        SQLX = SQLX & "0 AS TT_COEFLAG "
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP," & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 45
  '--------------------------------------------------- 11OK
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_SKILLD,TT_EXPFACT,TT_SKLDTE,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, SE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'11' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_SKILLD,"
    SQLX = SQLX & "SE_LEVEL AS TT_EXPFACT,"
    SQLX = SQLX & "SE_DATE AS TT_SKLDTE "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_EMPSKL as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.SE_SKILL = T1.TB_KEY) AND (Z.SE_SKILL_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE SE_EMPNBR IN " & xEmpList
    
    'Hemu
    SQLX = SQLX & " ORDER BY SE_DATE DESC "
    'Hemu
    
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '11' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_SKILLD,TT_EXPFACT,TT_SKLDTE,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'11' AS TT_RECNBR,"
        SQLX = SQLX & "'No Skill' AS TT_SKILLD,"
        SQLX = SQLX & "Null AS TT_EXPFACT,"
        SQLX = SQLX & "Null AS TT_SKLDTE "
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, " & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 50
    '--------------------------------------------------- 13OK
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_DEGREED,TT_YEAR,TT_MAJORD,TT_MINORD,TT_COMPL,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, EU_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'13' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_DEGREED,"
    SQLX = SQLX & "EU_YEAR AS TT_YEAR,"
    SQLX = SQLX & "T2.TB_DESC AS TT_MAJORD,"
    SQLX = SQLX & "T3.TB_DESC AS TT_MINORD, "
    SQLX = SQLX & "EU_COMP AS TT_COMPL"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (((Term_EDU as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.EU_DEGREE = T1.TB_KEY) AND (Z.EU_DEGREE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.EU_MAJOR = T2.TB_KEY) AND (Z.EU_MAJOR_TABL = T2.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T3 ON (Z.EU_MINOR = T3.TB_KEY) AND (Z.EU_MINOR_TABL = T3.TB_NAME))"
    SQLX = SQLX & " WHERE EU_EMPNBR IN " & xEmpList
    
    SQLX = SQLX & " ORDER BY EU_YEAR DESC"
    
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '13' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_DEGREED,TT_YEAR,TT_MAJORD,TT_MINORD,TT_COMPL,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'13' AS TT_RECNBR,"
        SQLX = SQLX & "'No Formal Education' AS TT_DEGREED,"
        SQLX = SQLX & "Null AS TT_YEAR,"
        SQLX = SQLX & "Null AS TT_MAJORD,"
        SQLX = SQLX & "Null AS TT_MINORD, "
        SQLX = SQLX & "Null AS TT_COMPL"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, " & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 55
  '--------------------------------------------------- 15OK
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_CTYPED,TT_COURSE,TT_DATCOMP,TT_RESULTD,TT_TBCO,TT_TBEMP,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, ES_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'15' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_CTYPED,"
    SQLX = SQLX & "ES_COURSE AS TT_COURSE,"
    SQLX = SQLX & "ES_DATCOMP AS TT_DATCOMP,"
    SQLX = SQLX & "T2.TB_DESC AS TT_RESULTD,"
    SQLX = SQLX & "ES_TBCO AS TT_TBCO,"
    SQLX = SQLX & "ES_TBEMP AS TT_TBEMP"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_HREDSEM as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.ES_CTYPE = T1.TB_KEY) AND (Z.ES_CTYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " LEFT JOIN HRTABL AS T2 ON (Z.ES_RESULTS = T2.TB_KEY) AND (Z.ES_RESULTS_TABL = T2.TB_NAME)"
    SQLX = SQLX & " WHERE ES_EMPNBR IN " & xEmpList
    SQLX = SQLX & " ORDER BY ES_DATCOMP DESC, ES_CTYPE ASC"
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '15' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_CTYPED,TT_COURSE,TT_DATCOMP,TT_RESULTD,TT_TBCO,TT_TBEMP,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'15' AS TT_RECNBR,"
        SQLX = SQLX & "'No Course/Seminar' AS TT_CTYPED,"
        SQLX = SQLX & "Null AS TT_COURSE,"
        SQLX = SQLX & "Null AS TT_DATCOMP,"
        SQLX = SQLX & "Null AS TT_RESULTD,"
        SQLX = SQLX & "Null AS TT_TBCO,"
        SQLX = SQLX & "Null AS TT_TBEMP"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, " & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 60
  '--------------------------------------------------- 17OK
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_CODED,TT_DUES,TT_COMPPD,TT_RENEWDT,TT_BEGINDT,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, TD_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'17' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_CODED,"
    SQLX = SQLX & "TD_DUES AS TT_DUES,"
    SQLX = SQLX & "TD_COMPPD AS TT_COMPPD,"
    SQLX = SQLX & "TD_RENEWDT AS TT_RENEWDT,"
    SQLX = SQLX & "TD_BEGINDT AS TT_BEGINDT"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_TRADE as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.TD_CODE = T1.TB_KEY) AND (Z.TD_CODE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE TD_EMPNBR IN " & xEmpList
    
    SQLX = SQLX & " ORDER BY TD_BEGINDT DESC"
    
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '17' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_CODED,TT_DUES,TT_COMPPD,TT_RENEWDT,TT_BEGINDT,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'17' AS TT_RECNBR,"
        SQLX = SQLX & "'No Association/Membership' AS TT_CODED,"
        SQLX = SQLX & "Null AS TT_DUES,"
        SQLX = SQLX & "Null AS TT_COMPPD,"
        SQLX = SQLX & "Null AS TT_RENEWDT,"
        SQLX = SQLX & "Null AS TT_BEGINDT"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP," & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 65
  '--------------------------------------------------- 19OK
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_BCODED,TT_COVER,TT_BEDATE,TT_BAMT,TT_CCOST,TT_ECOST,TT_PCE,TT_PCC,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, BF_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'19' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_BCODED,"
    SQLX = SQLX & "BF_COVER AS TT_COVER,"
    SQLX = SQLX & "BF_EDATE AS TT_BEDATE,"
    SQLX = SQLX & "BF_AMT AS TT_BAMT,"
    SQLX = SQLX & "BF_CCOST AS TT_CCOST,"
    SQLX = SQLX & "BF_ECOST AS TT_ECOST,"
    SQLX = SQLX & "BF_PCE AS TT_PCE,"
    SQLX = SQLX & "BF_PCC AS TT_PCC"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_HRBENFT as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.BF_BCODE = T1.TB_KEY) AND (Z.BF_BCODE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE BF_EMPNBR IN " & xEmpList
    SQLX = SQLX & " ORDER BY BF_BCODE, BF_EDATE DESC"
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '19' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_BCODED,TT_COVER,TT_BEDATE,TT_BAMT,TT_CCOST,TT_ECOST,TT_PCE,TT_PCC,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'19' AS TT_RECNBR,"
        SQLX = SQLX & "'No Benefit' AS TT_BCODED,"
        SQLX = SQLX & "Null AS TT_COVER,"
        SQLX = SQLX & "Null AS TT_BEDATE,"
        SQLX = SQLX & "Null AS TT_BAMT,"
        SQLX = SQLX & "Null AS TT_CCOST,"
        SQLX = SQLX & "Null AS TT_ECOST,"
        SQLX = SQLX & "Null AS TT_PCE,"
        SQLX = SQLX & "Null AS TT_PCC"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, " & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 70
  '--------------------------------------------------- 20
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_BCODED,TT_BNAME,TT_BRELATE,TT_BDOB,TT_PCE,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, BD_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'20' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_BCODED,"
    SQLX = SQLX & "BD_BNAME AS TT_BNAME,"
    SQLX = SQLX & "BD_RELATE AS TT_BRELATE,"
    SQLX = SQLX & "BD_DOB AS TT_BDOB,"
    SQLX = SQLX & "BD_PC AS TT_PCE"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_HRBENS as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.BD_BCODE = T1.TB_KEY) AND (Z.BD_BCODE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE BD_EMPNBR IN " & xEmpList
    SQLX = SQLX & " ORDER BY BD_BCODE"
    gdbAdoIhr001X.Execute SQLX
    
    'Ticket #13685, if no record and then add a record to show it record
    SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM Term_HREMP WHERE ED_EMPNBR IN " & xEmpList & " "
    SQLQ = SQLQ & "AND NOT (ED_EMPNBR IN (SELECT TT_EMPNBR FROM HRTERMEMPWRK WHERE TT_RECNBR = '20' AND TT_WRKEMP = '" & glbUserID & "')) "
    If rsTTemp.State <> 0 Then rsTTemp.Close
    rsTTemp.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    Do While Not rsTTemp.EOF
        SQLX = "INSERT INTO HRTERMEMPWRK "
        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
        SQLX = SQLX & "TT_BCODED,TT_BNAME,TT_BRELATE,TT_BDOB,TT_PCE,TT_WRKEMP,TERM_SEQ) "
        SQLX = SQLX & in_SQL(glbIHRDBW)
        
        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, " & rsTTemp("ED_EMPNBR") & " AS TT_EMPNBR,"
        SQLX = SQLX & "'20' AS TT_RECNBR,"
        SQLX = SQLX & "'No Beneficiary' AS TT_BCODED,"
        SQLX = SQLX & "'' AS TT_BNAME,"
        SQLX = SQLX & "'' AS TT_BRELATE,"
        SQLX = SQLX & "Null AS TT_BDOB,"
        SQLX = SQLX & "Null AS TT_PCE"
        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, " & rsTTemp("TERM_SEQ") & ""
        gdbAdoIhr001X.Execute SQLX
        rsTTemp.MoveNext
    Loop
    rsTTemp.Close
    
MDIMain.panHelp(0).FloodPercent = 75
  '--------------------------------------------------- 21
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_VAC,TT_SICK,TT_PVAC,TT_PSICK,TT_VACHRS,TT_SICKHRS,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, ED_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'21' AS TT_RECNBR,"
    SQLX = SQLX & "ED_VAC AS TT_VAC,"
    SQLX = SQLX & "ED_SICK AS TT_SICK,"
    SQLX = SQLX & "ED_PVAC AS TT_PVAC,"
    SQLX = SQLX & "ED_PSICK AS TT_PSICK,"
    SQLX = SQLX & "ED_VACT as TT_VACHRS,"
    SQLX = SQLX & "ED_SICKT as TT_SICKHRS"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP, TERM_SEQ "
    SQLX = SQLX & " FROM Term_HREMP "
    SQLX = SQLX & " WHERE ED_EMPNBR IN " & xEmpList
    gdbAdoIhr001X.Execute SQLX
    
MDIMain.panHelp(0).FloodPercent = 80

    SQLX = "UPDATE HRTERMEMPWRK SET "
    SQLX = SQLX & " HRTERMEMPWRK.TT_OTHRS =ATTHRS.OTHRS, "
    SQLX = SQLX & " HRTERMEMPWRK.TT_CTHRS = ATTHRS.CTHRS, "
    SQLX = SQLX & " HRTERMEMPWRK.TT_WCBHRS =ATTHRS.WCBHRS "
    SQLX = SQLX & " FROM HRTERMEMPWRK INNER JOIN ("
    SQLX = SQLX & " SELECT AD_EMPNBR, TERM_SEQ, "
    SQLX = SQLX & " SUM(CASE WHEN LEFT(AD_REASON,2)='OT' THEN AD_HRS ELSE 0 END) AS OTHRS,"
    SQLX = SQLX & " SUM(CASE WHEN LEFT(AD_REASON,2)='CT' THEN AD_HRS ELSE 0 END) AS CTHRS,"
    SQLX = SQLX & " SUM(CASE WHEN LEFT(AD_REASON,3)='WCB' THEN AD_HRS ELSE 0 END) AS WCBHRS "
    SQLX = SQLX & " FROM Term_ATTENDANCE "
    SQLX = SQLX & " WHERE AD_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND AD_DOA >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND AD_DOA <= " & Date_SQL(xDate2)
    SQLX = SQLX & " GROUP BY AD_EMPNBR, TERM_SEQ"
    SQLX = SQLX & " ) AS ATTHRS ON HRTERMEMPWRK.TT_EMPNBR=ATTHRS.AD_EMPNBR AND HRTERMEMPWRK.TERM_SEQ=ATTHRS.TERM_SEQ"
    SQLX = SQLX & " WHERE HRTERMEMPWRK.TT_RECNBR='21'"
    gdbAdoIhr001X.Execute SQLX
   
MDIMain.panHelp(0).FloodPercent = 85
    '-----------------------------------------------------22
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_ETYPE,TT_EFDATE,TT_ETDATE,TT_ENTITLE,TT_EACTUAL,TT_COEFLAG,TT_WRKEMP,TERM_SEQ) "
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, HE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'22' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_ETYPE,"
    SQLX = SQLX & "HE_FDATE AS TT_EFDATE,"
    SQLX = SQLX & "HE_TDATE AS TT_ETDATE,"
    SQLX = SQLX & "HE_ENTITLE AS TT_ENTITLE,"
    SQLX = SQLX & "HE_TAKEN AS TT_EACTUAL,"
    SQLX = SQLX & "HE_COE AS TT_COEFLAG"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_ENTHRS as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.HE_TYPE = T1.TB_KEY) AND (Z.HE_TYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE HE_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND HE_FDATE >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND HE_TDATE <= " & Date_SQL(xDate2)
    SQLX = SQLX & " ORDER BY HE_TYPE"
    gdbAdoIhr001X.Execute SQLX
    
MDIMain.panHelp(0).FloodPercent = 90
  '--------------------------------------------------- 23
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_DTYPE,TT_DFDATE,TT_DTDATE,TT_DENTITL,TT_DACTUAL,TT_COEFLAG,TT_WRKEMP,TERM_SEQ)"
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, DE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'23' AS TT_RECNBR,"
    SQLX = SQLX & "T1.TB_DESC AS TT_DTYPE,"
    SQLX = SQLX & "DE_FDATE AS TT_DFDATE,"
    SQLX = SQLX & "DE_TDATE AS TT_DTDATE,"
    SQLX = SQLX & "DE_ENTITLE AS TT_DENTITL,"
    SQLX = SQLX & "DE_ACTUAL AS TT_DACTUAL,"
    SQLX = SQLX & "DE_COST_OF_EMPLOYMENT AS TT_COEFLAG"
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
    SQLX = SQLX & " FROM (Term_DOLENT as Z "
    SQLX = SQLX & " LEFT JOIN HRTABL AS T1 ON (Z.DE_TYPE = T1.TB_KEY) AND (Z.DE_TYPE_TABL = T1.TB_NAME))"
    SQLX = SQLX & " WHERE DE_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND DE_FDATE >= " & Date_SQL(xDate1)
    SQLX = SQLX & " AND DE_TDATE <= " & Date_SQL(xDate2)
    SQLX = SQLX & " ORDER BY DE_FDATE DESC"
    gdbAdoIhr001X.Execute SQLX
    
MDIMain.panHelp(0).FloodPercent = 95
   '--------------------------------------------------- 25
    SQLX = "INSERT INTO HRTERMEMPWRK "
    SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
    SQLX = SQLX & "TT_OLDDEPT,TT_NEWDEPT,TT_OLDDIV,TT_NEWDIV,TT_OLDEMP,TT_NEWEMP,TT_OLDPT,TT_NEWPT,TT_OLDORG,TT_NEWORG,TT_CHGDATE,TT_WRKEMP,TT_TODATE,TERM_SEQ) "    'Hemu TT_TODATE
    SQLX = SQLX & in_SQL(glbIHRDBW)
    
    SQLX = SQLX & "SELECT '001' AS TT_COMPNO, EE_EMPNBR AS TT_EMPNBR,"
    SQLX = SQLX & "'25' AS TT_RECNBR,"
    SQLX = SQLX & "EE_OLDDEPT AS TT_OLDDEPT,"
    SQLX = SQLX & "EE_NEWDEPT AS TT_NEWDEPT,"
    SQLX = SQLX & "EE_OLDDIV AS TT_OLDDIV,"
    SQLX = SQLX & "EE_NEWDIV AS TT_NEWDIV,"
    SQLX = SQLX & "EE_OLDSTAT AS TT_OLDEMP,"
    SQLX = SQLX & "EE_NEWSTAT AS TT_NEWEMP,"
    SQLX = SQLX & "EE_OLDPT AS TT_OLDPT,"
    SQLX = SQLX & "EE_NEWPT AS TT_NEWPT,"
    SQLX = SQLX & "EE_OLDORG AS TT_OLDORG,"
    SQLX = SQLX & "EE_NEWORG AS TT_NEWORG,"
    SQLX = SQLX & "EE_CHGDATE AS TT_CHGDATE "
    SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
    
    'Hemu
    SQLX = SQLX & ",EE_TODATE AS TT_TODATE,TERM_SEQ "
    'Hemu
    
    SQLX = SQLX & " FROM Term_HREMPHIS "
    SQLX = SQLX & " WHERE EE_EMPNBR IN " & xEmpList
    SQLX = SQLX & " AND (EE_OLDDEPT IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWDEPT IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDDIV IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWDIV IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDSTAT IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWSTAT IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDORG IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWORG IS NOT NULL "
    SQLX = SQLX & " OR EE_OLDPT IS NOT NULL "
    SQLX = SQLX & " OR EE_NEWPT IS NOT NULL)"
    gdbAdoIhr001X.Execute SQLX
  '--------------------------------------------------- 26
' No Term table for HRSTATUS
'    If Not glbLinamar Then
'        SQLX = "INSERT INTO HRTERMEMPWRK "
'        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
'        SQLX = SQLX & "TT_OLDEMP,TT_NEWEMP,TT_EFDATE,TT_ETDATE,TT_JOBCODE,TT_EMP,TT_WRKEMP,TERM_SEQ) "
'        SQLX = SQLX & in_SQL(glbIHRDBW)
'
'        SQLX = SQLX & "SELECT '001' AS SC_COMPNO, SC_EMPNBR AS TT_EMPNBR,"
'        SQLX = SQLX & "'26' AS TT_RECNBR,"
'        SQLX = SQLX & "SC_OLDEMP AS TT_OLDEMP,"
'        SQLX = SQLX & "SC_NEWEMP AS TT_NEWEMP,"
'        SQLX = SQLX & "SC_FDATE AS TT_EFDATE, "
'        SQLX = SQLX & "SC_TDATE AS TT_ETDATE, "
'        SQLX = SQLX & "SC_JOB AS TT_JOBCODE,"
'        SQLX = SQLX & "SC_REASON AS TT_EMP,"
'        SQLX = SQLX & "'" & glbUserID & "' AS TT_WRKEMP,TERM_SEQ "
'        SQLX = SQLX & " FROM HRSTATUS "
'        SQLX = SQLX & " WHERE SC_EMPNBR IN " & xEmplist
'        SQLX = SQLX & " ORDER BY SC_LDATE,SC_LTIME"
'        gdbAdoIhr001.Execute SQLX
'    End If
     '------------------------------------------------ end
    
    '--------------------------------------------------- 27
'No Term tables for WFC Pension
'    'WFC Pension Beneficiary Information
'    If glbWFC Then
'        SQLX = "INSERT INTO HRTERMEMPWRK "
'        SQLX = SQLX & "(TT_COMPNO,TT_EMPNBR,TT_RECNBR,"
'        SQLX = SQLX & "TT_BCODED,TT_BNAME,TT_BRELATE,TT_BDOB,TT_WRKEMP,TERM_SEQ) "
'        SQLX = SQLX & in_SQL(glbIHRDBW)
'
'        SQLX = SQLX & "SELECT '001' AS TT_COMPNO, PE_EMPNBR AS TT_EMPNBR,"
'        SQLX = SQLX & "'27' AS TT_RECNBR,"
'        SQLX = SQLX & "PE_PENSIONTYPE AS TT_BCODED,"
'        SQLX = SQLX & "PE_BEN_NAME AS TT_BNAME,"
'        SQLX = SQLX & "PE_BEN_RELATE AS TT_BRELATE,"
'        SQLX = SQLX & "PE_BEN_DOB AS TT_BDOB"
'        SQLX = SQLX & ",'" & glbUserID & "' AS TT_WRKEMP "
'        SQLX = SQLX & " FROM HRP_PENSION_BENEFICIARY "
'        SQLX = SQLX & " WHERE PE_EMPNBR IN " & xEmplist
'        SQLX = SQLX & " ORDER BY PE_EMPNBR,PE_BEN_NAME"
'        gdbAdoIhr001.Execute SQLX
'    End If
'    '--------------------------------------------------- end

End Sub

Public Sub glbEmpWrkAccess(xEmpList, xDate1, xDate2)
Dim rsEmp As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim SQLQ
Dim xxx, xx1, xEmpNbr, xNoDay#, xDate0, xWCB, xOT, xCT
'SQLQ = "SELECT ED_COMPNO,ED_EMPNBR,ED_DOH,ED_DOB,ED_SENDTE,ED_LANG1,ED_LANG2," & _
       "ED_VAC,ED_SICK,ED_PVAC,ED_PSICK,ED_VACT,ED_SICKT " & _
       "FROM HREMP WHERE ED_EMPNBR IN " & xEmplist
SQLQ = "SELECT ED_COMPNO,ED_EMPNBR,ED_DOH,ED_DOB,ED_SENDTE," & _
       "ED_VAC,ED_SICK,ED_PVAC,ED_PSICK,ED_VACT,ED_SICKT " & _
       "FROM HREMP WHERE ED_EMPNBR IN " & xEmpList
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

rsWRK.Open "HREMPWRK", gdbAdoIhr001W, adOpenStatic, adLockPessimistic, adCmdTableDirect
rsEmp.MoveLast
xxx = rsEmp.RecordCount / 80
xx1 = 0
rsEmp.MoveFirst
Do Until rsEmp.EOF
    xx1 = xx1 + 1
    MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) + 10
    xEmpNbr = rsEmp("ED_EMPNBR")
    '--------------------------------------------------- 01
    'FName.Caption = "HREMP"
    rsWRK.AddNew
    rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
    rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
    rsWRK("TT_RECNBR") = "01"
    If IsDate(rsEmp("ED_DOH")) Then
        xNoDay# = DateDiff("d", rsEmp("ED_DOH"), Now) / 365
        rsWRK("TT_NUMERIC") = xNoDay#
    End If
    If IsDate(rsEmp("ED_DOB")) Then
        xNoDay# = DateDiff("d", rsEmp("ED_DOB"), Now) / 365
        rsWRK("TT_FTE") = xNoDay#
    End If
    If IsDate(rsEmp("ED_SENDTE")) Then
        xNoDay# = DateDiff("d", rsEmp("ED_SENDTE"), Now) / 365
        rsWRK("TT_FTEHRS") = xNoDay#
    End If
    rsWRK("TT_WRKEMP") = glbUserID
    rsWRK.Update
'    '--------------------------------------------------- 10
'    If Not IsNull(rsEmp("ED_LANG1")) Or Not IsNull(rsEmp("ED_LANG2")) Then
'        rsWRK.AddNew
'        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
'        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
'        rsWRK("TT_RECNBR") = "10"
'        If Not IsNull(rsEmp("ED_LANG1")) Then rsWRK("TT_LANG1") = READTABLE("EDL1", rsEmp("ED_LANG1"), True)
'        If Not IsNull(rsEmp("ED_LANG2")) Then rsWRK("TT_LANG2") = READTABLE("EDL1", rsEmp("ED_LANG2"), True)
'        rsWRK("TT_WRKEMP") = glbUserID
'        rsWRK.Update
'    End If
'George Modified on Mar 21,2006 #10574
    '--------------------------------------------------- 10
    rsTB.Open "SELECT * FROM HR_LANGUAGE WHERE EL_EMPNBR=" & xEmpNbr & " ORDER BY EL_LANGNO", gdbAdoIhr001, adOpenKeyset
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "10"
            rsWRK("TT_LANG1") = "No Language"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            If Not IsNull(rsTB("EL_LANG_SPOKEN")) Or Not IsNull(rsTB("EL_LANG_WRITTEN")) Then
                rsWRK.AddNew
                rsWRK("TT_COMPNO") = rsTB("EL_COMPNO")
                rsWRK("TT_EMPNBR") = rsTB("EL_EMPNBR")
                rsWRK("TT_RECNBR") = "10"
                If Not IsNull(rsTB("EL_LANG_SPOKEN")) Then rsWRK("TT_LANG1") = READTABLE("EDL1", rsTB("EL_LANG_SPOKEN"), True)
                If Not IsNull(rsTB("EL_LANG_WRITTEN")) Then rsWRK("TT_LANG2") = READTABLE("EDL1", rsTB("EL_LANG_WRITTEN"), True)
                rsWRK("TT_WRKEMP") = glbUserID
                rsWRK.Update
            End If
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
'George Modified on Mar 21,2006 #10574
    '--------------------------------------------------- 03
    'FName.Caption = "HRDEPEND"
    rsTB.Open "SELECT * FROM HRDEPEND WHERE DP_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenKeyset
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "03"
            rsWRK("TT_NAMEFLD") = "No Dependent"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "03"
            rsWRK("TT_SEX") = rsTB("DP_SEX")
            rsWRK("TT_NAMEFLD") = Trim(rsTB("DP_SNAME")) & ", " & Trim(rsTB("DP_FNAME"))
            rsWRK("TT_CHAR10") = rsTB("DP_RELATE")
            rsWRK("TT_DATEFLD") = rsTB("DP_DOB")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 05, 08
    'FName.Caption = "HR_SALARY_HISTORY"
    
    rsTB.Open "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpNbr & _
              " ORDER BY SH_CURRENT,SH_EDATE DESC,SH_LDATE,SH_ID", gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then xDate0 = rsTB("SH_SDATE")
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        If rsTB("SH_CURRENT") Then
            rsWRK("TT_RECNBR") = "05"
        Else
            rsWRK("TT_RECNBR") = "08"
            rsWRK("TT_COMPA") = rsTB("SH_COMPA")
        End If
        rsWRK("TT_JOB") = ReadJob(rsTB("SH_JOB"))
        rsWRK("TT_GRID") = READTABLE("JBGD", rsTB("SH_GRID"), True)
        'added by Bryan 22/Sep/05 Ticket#9343
        If glbCompSerial = "S/N - 2373W" Then 'Muskoka
            rsWRK("TT_SALARY") = rsTB("SH_TOTAL")
        Else
            rsWRK("TT_SALARY") = rsTB("SH_SALARY")
        End If
        rsWRK("TT_SALCD") = rsTB("SH_SALCD")
        rsWRK("TT_SEDATE") = rsTB("SH_EDATE")
        rsWRK("TT_SALR1") = READTABLE("SDRC", rsTB("SH_SREAS1"), True)
        rsWRK("TT_SALPC1") = rsTB("SH_SALPC1")
        'If Not IsNull(RSTB("SH_PCODE")) Then rswrk("TT_PCODE") = READTABLE("SDRC", RSTB("SH_PCODE"))
        rsWRK("TT_NEXTDAT") = rsTB("SH_NEXTDAT")
        rsWRK("TT_WHRS") = rsTB("SH_WHRS")
        rsWRK("TT_GRADE") = rsTB("SH_GRADE")
        If xDate0 <> rsTB("SH_SDATE") Then
            rsWRK("TT_DATECHG") = xDate0
            xDate0 = rsTB("SH_SDATE")
        End If
        If Not IsNull(rsTB("SH_SREAS2")) Then rsWRK("TT_SALR2") = READTABLE("SDRC", rsTB("SH_SREAS2"), True)
        If Not IsNull(rsTB("SH_SREAS3")) Then rsWRK("TT_SALR3") = READTABLE("SDRC", rsTB("SH_SREAS3"), True)
        rsWRK("TT_SALPC2") = rsTB("SH_SALPC2")
        rsWRK("TT_SALPC3") = rsTB("SH_SALPC3")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 04, 07
    'FName.Caption = "HR_JOB_HISTORY"
    rsTB.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & xEmpNbr & _
            " ORDER BY JH_SDATE DESC,JH_CURRENT,JH_ID", gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then xDate0 = rsTB("JH_SDATE")
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        If rsTB("JH_CURRENT") Then
            rsWRK("TT_RECNBR") = "04"
        Else
            rsWRK("TT_RECNBR") = "07"
        End If
        rsWRK("TT_JOB") = ReadJob(rsTB("JH_JOB"))
        rsWRK("TT_GRID") = READTABLE("JBGD", rsTB("JH_GRID"), True)
        rsWRK("TT_JREAS") = READTABLE("SDRC", rsTB("JH_JREASON"), True)
        rsWRK("TT_REPTAU") = rsTB("JH_REPTAU")
        If xDate0 <> rsTB("JH_SDATE") Then
            rsWRK("TT_DATECHG") = xDate0
            'xNoDay# = DateDiff("d", RSTB("JH_SDATE"), xDate0)
            'rswrk("TT_NUMERIC") = xNoDay#
            xDate0 = rsTB("JH_SDATE")
        End If
        rsWRK("TT_SDATE") = rsTB("JH_SDATE")
        rsWRK("TT_SHIFT") = rsTB("JH_SHIFT")
        rsWRK("TT_FTE") = rsTB("JH_FTENUM")
        rsWRK("TT_FTEHRS") = rsTB("JH_FTEHRS")
        rsWRK("TT_WHRS") = rsTB("JH_WHRS")
        rsWRK("TT_DHRS") = rsTB("JH_DHRS")
        rsWRK("TT_PHRS") = rsTB("JH_PHRS")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 06, 09
    'FName.Caption = "HR_PERFORM_HISTORY"
    rsTB.Open "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR=" & xEmpNbr & _
            " ORDER BY PH_CURRENT,PH_PREVIEW DESC,PH_ID", gdbAdoIhr001, adOpenKeyset
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        If rsTB("PH_CURRENT") Then
            rsWRK("TT_RECNBR") = "06"
        Else
            rsWRK("TT_RECNBR") = "09"
        End If
        rsWRK("TT_JOB") = ReadJob(rsTB("PH_JOB"))
        rsWRK("TT_REPTAU") = rsTB("PH_REPTAU")
        If Not IsNull(rsTB("PH_PCODE")) Then rsWRK("TT_PCODE") = READTABLE("SDPC", rsTB("PH_PCODE"), True)
        rsWRK("TT_PREVIEW") = rsTB("PH_PREVIEW")
        rsWRK("TT_PNEXT") = rsTB("PH_PNEXT")
        'rswrk("TT_PCOMM") = RSTB("PH_COMMENTS")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 12
    'FName.Caption = "HREARN"
    rsTB.Open "SELECT * FROM HREARN WHERE EMPNBR=" & xEmpNbr & _
              " AND FDATE>=" & Date_SQL(xDate1) & _
              " AND TDATE<=" & Date_SQL(xDate2) & _
              " ORDER BY EARN_TYPE,TDATE DESC ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "Empl/Code"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "12"
            rsWRK("TT_ETYPE") = "No Other Earnings"
            rsWRK("TT_COEFLAG") = 0
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "12"
            rsWRK("TT_ETYPE") = READTABLE("EARN", rsTB("EARN_TYPE"), True)
            rsWRK("TT_EFDATE") = rsTB("FDATE")
            rsWRK("TT_ETDATE") = rsTB("TDATE")
            rsWRK("TT_EACTUAL") = rsTB("ACT_DOLLAR")
            rsWRK("TT_COEFLAG") = rsTB("COST_OF_EMPLOYMENT")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 11
    'FName.Caption = "HREMPSKL"
    
    rsTB.Open "SELECT * FROM HREMPSKL WHERE SE_EMPNBR=" & xEmpNbr & " ORDER BY SE_DATE DESC", gdbAdoIhr001, adOpenKeyset
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "11"
            rsWRK("TT_SKILLD") = "No Skill"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "11"
            rsWRK("TT_SKILLD") = READTABLE("EDSK", rsTB("SE_SKILL"), True)
            rsWRK("TT_EXPFACT") = rsTB("SE_LEVEL")
            rsWRK("TT_SKLDTE") = rsTB("SE_DATE")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 13
    'FName.Caption = "HREDU"
    rsTB.Open "SELECT * FROM HREDU WHERE EU_EMPNBR=" & xEmpNbr & _
              " ORDER BY EU_SCHOOL,EU_YEAR DESC ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL/CODE"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "13"
            rsWRK("TT_DEGREED") = "No Formal Education"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "13"
            rsWRK("TT_DEGREED") = READTABLE("EUDE", rsTB("EU_DEGREE"), False)
            rsWRK("TT_YEAR") = rsTB("EU_YEAR")
            rsWRK("TT_MAJORD") = READTABLE("EUMJ", rsTB("EU_MAJOR"), False)
            rsWRK("TT_MINORD") = READTABLE("EUMJ", rsTB("EU_MINOR"), False)
            rsWRK("TT_COMPL") = rsTB("EU_COMP")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 15
    'FName.Caption = "HREDSEM"
    rsTB.Open "SELECT * FROM HREDSEM WHERE ES_EMPNBR=" & xEmpNbr & _
            " ORDER BY ES_CTYPE,ES_DATCOMP DESC", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL/CODE"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "15"
            rsWRK("TT_CTYPED") = "No Course/Seminar"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "15"
            rsWRK("TT_CTYPED") = READTABLE("ESCT", rsTB("ES_CTYPE"), False)
            rsWRK("TT_COURSE") = rsTB("ES_COURSE")
            rsWRK("TT_DATCOMP") = rsTB("ES_DATCOMP")
            rsWRK("TT_RESULTD") = READTABLE("ESRT", rsTB("ES_RESULTS"), False)
            rsWRK("TT_TBCO") = rsTB("ES_TBCO")
            rsWRK("TT_TBEMP") = rsTB("ES_TBEMP")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 14
    'FName.Caption = "HRTRADE"
    rsTB.Open "SELECT * FROM HRTRADE WHERE TD_EMPNBR=" & xEmpNbr & _
            " ORDER BY TD_CODE,TD_BEGINDT DESC", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL#/CODE"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "17"
            rsWRK("TT_CODED") = "No Association/Memberships"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "17"
            rsWRK("TT_CODED") = READTABLE("TDCD", rsTB("TD_CODE"), True)
            rsWRK("TT_DUES") = rsTB("TD_DUES")
            rsWRK("TT_COMPPD") = rsTB("TD_COMPPD")
            rsWRK("TT_RENEWDT") = rsTB("TD_RENEWDT")
            rsWRK("TT_BEGINDT") = rsTB("TD_BEGINDT")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 15
    'FName.Caption = "HRBENFT"
    rsTB.Open "SELECT * FROM HRBENFT WHERE BF_EMPNBR=" & xEmpNbr & _
            " ORDER BY BF_BCODE,BF_EDATE DESC,BF_BENE_ID", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL/CODE"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "19"
            rsWRK("TT_BCODED") = "No Benefit"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "19"
            rsWRK("TT_BCODED") = READTABLE("BNCD", rsTB("BF_BCODE"), True)
            rsWRK("TT_COVER") = rsTB("BF_COVER")
            rsWRK("TT_BEDATE") = rsTB("BF_EDATE")
            rsWRK("TT_BAMT") = rsTB("BF_AMT")
            rsWRK("TT_CCOST") = rsTB("BF_CCOST")
            rsWRK("TT_ECOST") = rsTB("BF_ECOST")
            rsWRK("TT_PCE") = rsTB("BF_PCE")
            rsWRK("TT_PCC") = rsTB("BF_PCC")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 21
    'FName.Caption = "HRBENS"
    rsTB.Open "SELECT * FROM HRBENS WHERE BD_EMPNBR=" & xEmpNbr & _
            " ORDER BY BD_BCODE,BD_ID", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "PrimaryKey"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "20"
            rsWRK("TT_BCODED") = "No Beneficiary"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "20"
            rsWRK("TT_BCODED") = READTABLE("BNCD", rsTB("BD_BCODE"), True)
            rsWRK("TT_BNAME") = rsTB("BD_BNAME")
            rsWRK("TT_BRELATE") = rsTB("BD_RELATE")
            rsWRK("TT_BDOB") = rsTB("BD_DOB")
            rsWRK("TT_PCE") = rsTB("BD_PC")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 21
    xWCB = 0
    xOT = 0
    xCT = 0
    'FName.Caption = "HR_ATTENDANCE"
    rsTB.Open "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNbr & _
              " AND AD_DOA>=CVDATE('" & xDate1 & "') " & _
              " AND AD_DOA<=CVDATE('" & xDate2 & "') " & _
              " AND (LEFT(AD_REASON,3)='WCB' OR LEFT(AD_REASON,2)='CT' OR LEFT(AD_REASON,2)='OT') " & _
              " ORDER BY AD_REASON,AD_DOA DESC", gdbAdoIhr001, adOpenKeyset
    
    'RSTB.Index = "EMPL/CODE"
    Do Until rsTB.EOF
        If Left(rsTB("AD_REASON"), 3) = "WCB" Then xWCB = xWCB + rsTB("AD_HRS")                                     '
        If Left(rsTB("AD_REASON"), 2) = "CT" Then xCT = xCT + rsTB("AD_HRS")                                        '
        If Left(rsTB("AD_REASON"), 2) = "OT" Then xOT = xOT + rsTB("AD_HRS")                                        '
        rsTB.MoveNext                                                                                                 '
        If rsTB.EOF Then Exit Do                                                                                      '
    Loop                                                                                                            '
    rsWRK.AddNew
    rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
    rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
    rsWRK("TT_RECNBR") = "21"
    rsWRK("TT_VAC") = rsEmp("ED_VAC")
    rsWRK("TT_SICK") = rsEmp("ED_SICK")
    rsWRK("TT_PVAC") = rsEmp("ED_PVAC")
    rsWRK("TT_PSICK") = rsEmp("ED_PSICK")
    rsWRK("TT_OTHRS") = xOT
    rsWRK("TT_CTHRS") = xCT
    rsWRK("TT_WCBHRS") = xWCB
    rsWRK("TT_VACHRS") = rsEmp("ED_VACT")
    rsWRK("TT_SICKHRS") = rsEmp("ED_SICKT")
    rsWRK("TT_WRKEMP") = glbUserID
    rsWRK.Update
    rsTB.Close
    '--------------------------------------------------22
    'FName.Caption = "HRENTHRS"
    rsTB.Open "SELECT * FROM HRENTHRS WHERE HE_EMPNBR=" & xEmpNbr & _
              " AND HE_FDATE>=" & Date_SQL(xDate1) & _
              " AND HE_FDATE<=" & Date_SQL(xDate2) & _
              " ORDER BY HE_TYPE,HE_TDATE DESC", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "PrimaryKey"
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        rsWRK("TT_RECNBR") = "22"
        rsWRK("TT_ETYPE") = READTABLE("ADRE", rsTB("HE_TYPE"), True)
        rsWRK("TT_EFDATE") = rsTB("HE_FDATE")
        rsWRK("TT_ETDATE") = rsTB("HE_TDATE")
        rsWRK("TT_ENTITLE") = rsTB("HE_ENTITLE")
        rsWRK("TT_EACTUAL") = rsTB("HE_TAKEN")
        rsWRK("TT_COEFLAG") = rsTB("HE_COE")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 23
    'FName.Caption = "HRDOLENT"
    rsTB.Open "SELECT * FROM HRDOLENT WHERE DE_EMPNBR=" & xEmpNbr & _
              " AND DE_FDATE>=" & Date_SQL(xDate1) & _
              " AND DE_FDATE<=" & Date_SQL(xDate2) & _
              " ORDER BY DE_TYPE,DE_TDATE DESC ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL/CODE"
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        rsWRK("TT_RECNBR") = "23"
        rsWRK("TT_DTYPE") = READTABLE("EDOL", rsTB("DE_TYPE"), True)
        rsWRK("TT_DFDATE") = rsTB("DE_FDATE")
        rsWRK("TT_DTDATE") = rsTB("DE_TDATE")
        rsWRK("TT_DENTITL") = rsTB("DE_ENTITLE")
        rsWRK("TT_DACTUAL") = rsTB("DE_ACTUAL")
        rsWRK("TT_COEFLAG") = rsTB("DE_COST_OF_EMPLOYMENT")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 25
    'FName.Caption = "HREMPHIS"
    rsTB.Open "SELECT * FROM HREMPHIS WHERE EE_EMPNBR=" & xEmpNbr & _
              " ORDER BY EE_CHGDATE ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPNBR"
    Do Until rsTB.EOF
        If Not IsNull(rsTB("EE_OLDDEPT")) Or Not IsNull(rsTB("EE_NEWDEPT")) Or Not IsNull(rsTB("EE_OLDDIV")) Or Not IsNull(rsTB("EE_NEWDIV")) Or Not IsNull(rsTB("EE_OLDSTAT")) Or Not IsNull(rsTB("EE_NEWSTAT")) Or Not IsNull(rsTB("EE_OLDORG")) Or Not IsNull(rsTB("EE_NEWORG")) Or Not IsNull(rsTB("EE_OLDPT")) Or Not IsNull(rsTB("EE_NEWPT")) Then
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "25"
            rsWRK("TT_OLDDEPT") = rsTB("EE_OLDDEPT")
            rsWRK("TT_NEWDEPT") = rsTB("EE_NEWDEPT")
            rsWRK("TT_OLDDIV") = rsTB("EE_OLDDIV")
            rsWRK("TT_NEWDIV") = rsTB("EE_NEWDIV")
            rsWRK("TT_OLDEMP") = rsTB("EE_OLDSTAT")
            rsWRK("TT_NEWEMP") = rsTB("EE_NEWSTAT")
            'rswrk("TT_OLDJOB") = RSTB("EE_OLDJOB")
            'rswrk("TT_NEWJOB") = RSTB("EE_NEWJOB")
            rsWRK("TT_OLDPT") = rsTB("EE_OLDPT")
            rsWRK("TT_NEWPT") = rsTB("EE_NEWPT")
            rsWRK("TT_OLDORG") = rsTB("EE_OLDORG")
            rsWRK("TT_NEWORG") = rsTB("EE_NEWORG")
            'rswrk("TT_OLDGL") = RSTB("EE_OLDGLNO")
            'rswrk("TT_NEWGL") = RSTB("EE_NEWGLNO")
            'rswrk("TT_FLAG") = RSTB("EE_DOTFLAG")
            rsWRK("TT_CHGDATE") = rsTB("EE_CHGDATE")
            rsWRK("TT_WRKEMP") = glbUserID
            'Hemu
            rsWRK("TT_TODATE") = rsTB("EE_TODATE")
            'Hemu
            rsWRK.Update
        End If
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 26
    'FName.Caption = "HRSTATUS"
    rsTB.Open "SELECT * FROM HRSTATUS WHERE SC_EMPNBR=" & xEmpNbr & _
              " ORDER BY SC_LDATE,SC_LTIME ", gdbAdoIhr001, adOpenKeyset
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        rsWRK("TT_RECNBR") = "26"
        rsWRK("TT_OLDEMP") = rsTB("SC_OLDEMP")
        rsWRK("TT_NEWEMP") = rsTB("SC_NEWEMP")
        rsWRK("TT_EFDATE") = rsTB("SC_FDATE")
        rsWRK("TT_ETDATE") = rsTB("SC_TDATE")
        rsWRK("TT_JOBCODE") = rsTB("SC_JOB")
        rsWRK("TT_EMP") = rsTB("SC_REASON")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '------------------------------------------------ end
    rsEmp.MoveNext
    If rsEmp.EOF Then Exit Do
Loop
rsEmp.Close
rsWRK.Close
End Sub
Public Sub CreateVacBrant()
'Dim LgdbIhr001B As Database
'Dim tdfHRVac As TableDef
Dim fldTemp As Field
Dim SQLQ, Counter

    'Set LgdbIhr001B = OpenDatabase(glbIHRDBB, False, False, ";pwd=petman")
    'For Counter = 0 To LgdbIhr001B.TableDefs.count - 1
    '    If LgdbIhr001B.TableDefs(Counter).Name = "HRVacBrant" Then
    '        GoTo ToEnd1B
    '        Exit For
    '    End If
    'Next Counter
    Exit Sub 'Ticket #23810 Franks 06/17/2013 added this table into their SQL database
    
    'SQLQ = "CREATE TABLE HRVacBrant (ED_COMPNO CHAR(3),ED_EMPNBR LONG,ED_PVAC DOUBLE,ED_VAC DOUBLE,ED_VACT DOUBLE, "
    'SQLQ = SQLQ & "ED_EFDATE DATE,ED_ETDATE DATE,ED_REDATE DATE)"
    'On Error Resume Next
    ''LgdbIhr001B.Execute SQLQ
    'gdbAdoIhr001B.Execute SQLQ
    'On Error GoTo 0
ToEnd1B:

    'LgdbIhr001B.Close
End Sub
Public Sub CreateHRRSP()
Dim LgdbIhr001 As Database, LgdbIhr001X As Database
Dim tdfHRRSP As TableDef
Dim fldTemp As Field
Dim SQLQ, Counter

    Set LgdbIhr001 = OpenDatabase(glbIHRDB, False, False, ";pwd=petman")
    For Counter = 0 To LgdbIhr001.TableDefs.count - 1
        If LgdbIhr001.TableDefs(Counter).name = "HRRSP" Then
            GoTo ToEnd1
            Exit For
        End If
    Next Counter

    SQLQ = "CREATE TABLE HRRSP (RS_COMPNO CHAR(3),RS_EMPNBR LONG,RS_PLAN_TABL CHAR(4),RS_PLAN CHAR(4) )"
    
    On Error Resume Next
    LgdbIhr001.Execute SQLQ
    SQLQ = "CREATE INDEX EMPNBR ON HRRSP ([RS_EMPNBR]);"
    LgdbIhr001.Execute SQLQ
    On Error GoTo 0
    Set LgdbIhr001 = OpenDatabase(glbIHRDB, False, False, ";pwd=petman")
    Set tdfHRRSP = LgdbIhr001.TableDefs("HRRSP")
    tdfHRRSP.Fields!RS_COMPNO.DefaultValue = Chr(34) & "001" & Chr(34)
    tdfHRRSP.Fields!RS_COMPNO.AllowZeroLength = True

    tdfHRRSP.Fields!RS_PLAN_TABL.DefaultValue = Chr(34) & "ERSP" & Chr(34)
    tdfHRRSP.Fields!RS_PLAN_TABL.AllowZeroLength = True

    tdfHRRSP.Fields!RS_PLAN.AllowZeroLength = True

    Set fldTemp = tdfHRRSP.CreateField("RS_YERPER", dbDouble)
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_YEEPER", dbDouble)
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_YERDOL", dbCurrency)
    tdfHRRSP.Fields.Append fldTemp
        
    Set fldTemp = tdfHRRSP.CreateField("RS_YEEDOL", dbCurrency)
    tdfHRRSP.Fields.Append fldTemp
        
    Set fldTemp = tdfHRRSP.CreateField("RS_CONDATE", dbDate)
    tdfHRRSP.Fields.Append fldTemp
            
    Set fldTemp = tdfHRRSP.CreateField("RS_COE", dbBoolean)
    tdfHRRSP.Fields.Append fldTemp
        
    Set fldTemp = tdfHRRSP.CreateField("RS_LDATE", dbDate)
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_LTIME", dbText, 8)
    fldTemp.AllowZeroLength = True
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_LUSER", dbDouble)
    tdfHRRSP.Fields.Append fldTemp
ToEnd1:
'------ For terminated employees

    Set LgdbIhr001X = OpenDatabase(glbIHRAUDIT, False, False, ";pwd=petman")
    For Counter = 0 To LgdbIhr001X.TableDefs.count - 1
        If LgdbIhr001X.TableDefs(Counter).name = "Term_HRRSP" Then
            GoTo ToEnd2
            Exit For
        End If
    Next Counter

    SQLQ = "CREATE TABLE Term_HRRSP (RS_COMPNO CHAR(3),RS_EMPNBR LONG,RS_PLAN_TABL CHAR(4),RS_PLAN CHAR(4) )"
    
    On Error Resume Next
    LgdbIhr001X.Execute SQLQ
    SQLQ = "CREATE INDEX EMPNBR ON Term_HRRSP ([RS_EMPNBR]);"
    LgdbIhr001X.Execute SQLQ
    On Error GoTo 0
    Set LgdbIhr001 = OpenDatabase(glbIHRAUDIT, False, False, ";pwd=petman")
    Set tdfHRRSP = LgdbIhr001.TableDefs("Term_HRRSP")
    tdfHRRSP.Fields!RS_COMPNO.DefaultValue = Chr(34) & "001" & Chr(34)
    tdfHRRSP.Fields!RS_COMPNO.AllowZeroLength = True

    tdfHRRSP.Fields!RS_PLAN_TABL.DefaultValue = Chr(34) & "ERSP" & Chr(34)
    tdfHRRSP.Fields!RS_PLAN_TABL.AllowZeroLength = True

    tdfHRRSP.Fields!RS_PLAN.AllowZeroLength = True
    
    Set fldTemp = tdfHRRSP.CreateField("RS_YERPER", dbDouble)
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_YEEPER", dbDouble)
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_YERDOL", dbCurrency)
    tdfHRRSP.Fields.Append fldTemp
        
    Set fldTemp = tdfHRRSP.CreateField("RS_YEEDOL", dbCurrency)
    tdfHRRSP.Fields.Append fldTemp
        
    Set fldTemp = tdfHRRSP.CreateField("RS_CONDATE", dbDate)
    tdfHRRSP.Fields.Append fldTemp
            
    Set fldTemp = tdfHRRSP.CreateField("RS_COE", dbBoolean)
    tdfHRRSP.Fields.Append fldTemp
        
    Set fldTemp = tdfHRRSP.CreateField("RS_LDATE", dbDate)
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_LTIME", dbText, 8)
    fldTemp.AllowZeroLength = True
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("RS_LUSER", dbDouble)
    tdfHRRSP.Fields.Append fldTemp
    
    Set fldTemp = tdfHRRSP.CreateField("TERM_SEQ", dbDouble)
    tdfHRRSP.Fields.Append fldTemp
    
    SQLQ = "CREATE INDEX TERM_SEQ ON Term_HRRSP ([TERM_SEQ]);"
    LgdbIhr001X.Execute SQLQ
    
ToEnd2:
    LgdbIhr001.Close
    LgdbIhr001X.Close
End Sub

Private Function READTABLE(Iname, Ikey, IOpt)
Dim SQLQ, rsTABL As New ADODB.Recordset
If IOpt Then READTABLE = "No Table Description" Else READTABLE = " "
If Ikey = "" Then Exit Function
SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = '" & Iname & "' and TB_KEY = '" & Ikey & "'"
rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
If Not rsTABL.EOF Then READTABLE = rsTABL("TB_DESC")
rsTABL.Close
End Function

Private Function ReadJob(IJob)
Dim SQLQ, rsJOB As New ADODB.Recordset
If IJob = "" Then Exit Function
ReadJob = "NO POSITION DESC - " & IJob
SQLQ = "SELECT JB_DESCR FROM HRJOB WHERE JB_CODE = '" & IJob & "'"
rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockReadOnly
If Not rsJOB.EOF Then ReadJob = rsJOB("JB_DESCR")
rsJOB.Close
End Function

Sub ChangeOtherEarnAmount(xEmpNo, xSalary, xType, xFDate, xTDate)
Dim RsHREARN As New ADODB.Recordset
Dim SQLQ
    'For New Salary
    If xType = "A" Then
        RsHREARN.Open "HREARN", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
        RsHREARN.AddNew '"BRVM"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "BRVM"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"EDUC"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "EDUC"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"JURY"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "JURY"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"LOA"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "LOA"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"OT1"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "OT1"
        RsHREARN("ACT_DOLLAR") = Val(xSalary) * 1.5
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"OT2"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "OT2"
        RsHREARN("ACT_DOLLAR") = Val(xSalary) * 2
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"REG"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "REG"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"RETI"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "RETI"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"S100"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "S100"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        'RsHREARN.AddNew '"SCONT"
        'RsHREARN("EMPNBR") = xEmpNo
        'RsHREARN("FDATE") = CVDate(xFDate)
        'RsHREARN("TDATE") = CVDate(xTDate)
        'RsHREARN("EARN_TYPE") = "SCONT"
        'RsHREARN("ACT_DOLLAR") = Val(xSalary)
        'RsHREARN("LDATE") = DATE
        'RsHREARN("LTIME") = Time$
        'RsHREARN("LUSER") = glbLEE_ID
        'RsHREARN.Update
        RsHREARN.AddNew '"SFLO"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "SFLO"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"SOFF"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "SOFF"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"ST1"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "ST1"
        RsHREARN("ACT_DOLLAR") = Val(xSalary) * 1.5
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"ST2"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "ST2"
        RsHREARN("ACT_DOLLAR") = Val(xSalary) * 2
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.AddNew '"WCB"
        RsHREARN("EMPNBR") = xEmpNo
        RsHREARN("FDATE") = CVDate(xFDate)
        RsHREARN("TDATE") = CVDate(xTDate)
        RsHREARN("EARN_TYPE") = "WCB"
        RsHREARN("ACT_DOLLAR") = Val(xSalary)
        RsHREARN("LDATE") = Date
        RsHREARN("LTIME") = Time$
        RsHREARN("LUSER") = glbLEE_ID
        RsHREARN.Update
        RsHREARN.Close
        
    End If
    
    'For Salary Change
    If xType = "M" Then
        Dim xValue
        SQLQ = "SELECT * FROM HREARN WHERE EMPNBR = " & xEmpNo
        RsHREARN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not RsHREARN.EOF
            xValue = 0
            Select Case RsHREARN("EARN_TYPE")
            Case "BRVM"
                xValue = Val(xSalary)
            Case "EDUC"
                xValue = Val(xSalary)
            Case "JURY"
                xValue = Val(xSalary)
            Case "LOA"
                xValue = Val(xSalary)
            Case "OT1"
                xValue = Val(xSalary) * 1.5
            Case "OT2"
                xValue = Val(xSalary) * 2
            Case "REG"
                xValue = Val(xSalary)
            Case "RETI"
                xValue = Val(xSalary)
            Case "S100"
                xValue = Val(xSalary)
            'Case "SCONT"
            '    xValue = Val(xSalary)
            Case "SFLO"
                xValue = Val(xSalary)
            Case "SOFF"
                xValue = Val(xSalary)
            Case "ST1"
                xValue = Val(xSalary) * 1.5
            Case "ST2"
                xValue = Val(xSalary) * 2
            Case "WCB"
                xValue = Val(xSalary)
            End Select
            If xValue <> 0 Then
                RsHREARN("ACT_DOLLAR") = xValue
                RsHREARN.Update
            End If
            RsHREARN.MoveNext
        Loop
        RsHREARN.Close
    End If
End Sub
' sam add for oracle database

Public Sub glbEmpWrkOracle(xEmpList, xDate1, xDate2)
Dim rsEmp As New ADODB.Recordset
Dim rsWRK As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim SQLQ As String
Dim xxx, xx1, xEmpNbr, xNoDay#, xDate0, xWCB, xOT, xCT
On Error GoTo ERR_EmpWrkOracle

'SQLQ = "SELECT ED_COMPNO,ED_EMPNBR,ED_DOH,ED_DOB,ED_SENDTE,ED_LANG1,ED_LANG2," & _
       "ED_VAC,ED_SICK,ED_PVAC,ED_PSICK,ED_VACT,ED_SICKT " & _
       "FROM HREMP WHERE ED_EMPNBR IN " & xEmplist
SQLQ = "SELECT ED_COMPNO,ED_EMPNBR,ED_DOH,ED_DOB,ED_SENDTE," & _
       "ED_VAC,ED_SICK,ED_PVAC,ED_PSICK,ED_VACT,ED_SICKT " & _
       "FROM HREMP WHERE ED_EMPNBR IN " & xEmpList
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic

rsWRK.Open "HREMPWRK", gdbAdoIhr001W, adOpenStatic, adLockPessimistic, adCmdTableDirect
rsEmp.MoveLast
xxx = rsEmp.RecordCount / 80
xx1 = 0
rsEmp.MoveFirst
Do Until rsEmp.EOF
    xx1 = xx1 + 1
    MDIMain.panHelp(0).FloodPercent = (xx1 / xxx) + 10
    xEmpNbr = rsEmp("ED_EMPNBR")
    '--------------------------------------------------- 01
    'FName.Caption = "HREMP"
    rsWRK.AddNew
    rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
    rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
    rsWRK("TT_RECNBR") = "01"
    If IsDate(rsEmp("ED_DOH")) Then
        xNoDay# = DateDiff("d", rsEmp("ED_DOH"), Now) / 365
        rsWRK("TT_NUMERIC") = xNoDay#
    End If
    If IsDate(rsEmp("ED_DOB")) Then
        xNoDay# = DateDiff("d", rsEmp("ED_DOB"), Now) / 365
        rsWRK("TT_FTE") = xNoDay#
    End If
    If IsDate(rsEmp("ED_SENDTE")) Then
        xNoDay# = DateDiff("d", rsEmp("ED_SENDTE"), Now) / 365
        rsWRK("TT_FTEHRS") = xNoDay#
    End If
    rsWRK("TT_WRKEMP") = glbUserID
    rsWRK.Update
    '--------------------------------------------------- 10
'    If Not IsNull(rsEmp("ED_LANG1")) Or Not IsNull(rsEmp("ED_LANG2")) Then
'        rsWRK.AddNew
'        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
'        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
'        rsWRK("TT_RECNBR") = "10"
'        If Not IsNull(rsEmp("ED_LANG1")) Then rsWRK("TT_LANG1") = READTABLE("EDL1", rsEmp("ED_LANG1"), True)
'        If Not IsNull(rsEmp("ED_LANG2")) Then rsWRK("TT_LANG2") = READTABLE("EDL1", rsEmp("ED_LANG2"), True)
'        rsWRK("TT_WRKEMP") = glbUserID
'        rsWRK.Update
'    End If

'George Modified on Mar 21,2006 #10574
    '--------------------------------------------------- 10
    rsTB.Open "SELECT * FROM HR_LANGUAGE WHERE EL_EMPNBR=" & xEmpNbr & " ORDER BY EL_LANGNO ASC", gdbAdoIhr001, adOpenKeyset
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "10"
            rsWRK("TT_LANG1") = "No Language"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            If Not IsNull(rsTB("EL_LANG_SPOKEN")) Or Not IsNull(rsTB("EL_LANG_WRITTEN")) Then
                rsWRK.AddNew
                rsWRK("TT_COMPNO") = rsTB("EL_COMPNO")
                rsWRK("TT_EMPNBR") = rsTB("EL_EMPNBR")
                rsWRK("TT_RECNBR") = "10"
                If Not IsNull(rsTB("EL_LANG_SPOKEN")) Then rsWRK("TT_LANG1") = READTABLE("EDL1", rsTB("EL_LANG_SPOKEN"), True)
                If Not IsNull(rsTB("EL_LANG_WRITTEN")) Then rsWRK("TT_LANG2") = READTABLE("EDL1", rsTB("EL_LANG_WRITTEN"), True)
                rsWRK("TT_WRKEMP") = glbUserID
                rsWRK.Update
            End If
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
'George Modified on Mar 21,2006 #10574
    '--------------------------------------------------- 03
    'FName.Caption = "HRDEPEND"
    rsTB.Open "SELECT * FROM HRDEPEND WHERE DP_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenKeyset
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "03"
            rsWRK("TT_NAMEFLD") = "No Dependent"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "03"
            rsWRK("TT_SEX") = rsTB("DP_SEX")
            rsWRK("TT_NAMEFLD") = Trim(rsTB("DP_SNAME")) & ", " & Trim(rsTB("DP_FNAME"))
            rsWRK("TT_CHAR10") = rsTB("DP_RELATE")
            rsWRK("TT_DATEFLD") = rsTB("DP_DOB")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 05, 08
    'FName.Caption = "HR_SALARY_HISTORY"
    


    rsTB.Open "SELECT * FROM HR_SALARY_HISTORY WHERE SH_EMPNBR = " & xEmpNbr & _
                  " ORDER BY SH_CURRENT,SH_EDATE DESC,SH_LDATE,SH_ID", gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then xDate0 = rsTB("SH_SDATE")
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        If rsTB("SH_CURRENT") Then
            rsWRK("TT_RECNBR") = "05"
        Else
            rsWRK("TT_RECNBR") = "08"
            rsWRK("TT_COMPA") = rsTB("SH_COMPA")
        End If
        rsWRK("TT_JOB") = ReadJob(rsTB("SH_JOB"))
        rsWRK("TT_GRID") = READTABLE("JBGD", rsTB("SH_GRID"), True)
        'added by Bryan 22/Sep/05 Ticket#9343
        If glbCompSerial = "S/N - 2373W" Then 'Muskoka
            rsWRK("TT_SALARY") = rsTB("SH_TOTAL")
        Else
            rsWRK("TT_SALARY") = rsTB("SH_SALARY")
        End If
        rsWRK("TT_SALCD") = rsTB("SH_SALCD")
        rsWRK("TT_SEDATE") = rsTB("SH_EDATE")
        rsWRK("TT_SALR1") = READTABLE("SDRC", rsTB("SH_SREAS1"), True)
        rsWRK("TT_SALPC1") = rsTB("SH_SALPC1")
        'If Not IsNull(RSTB("SH_PCODE")) Then rswrk("TT_PCODE") = READTABLE("SDRC", RSTB("SH_PCODE"))
        rsWRK("TT_NEXTDAT") = rsTB("SH_NEXTDAT")
        rsWRK("TT_WHRS") = rsTB("SH_WHRS")
        
        If glbCompSerial = "S/N - 2351W" Then 'Burlington Tech
            rsWRK("TT_GRADE") = "00"
        Else
            rsWRK("TT_GRADE") = rsTB("SH_GRADE")
        End If
        
        If xDate0 <> rsTB("SH_SDATE") Then
            rsWRK("TT_DATECHG") = xDate0
            xDate0 = rsTB("SH_SDATE")
        End If
        If Not IsNull(rsTB("SH_SREAS2")) Then rsWRK("TT_SALR2") = READTABLE("SDRC", rsTB("SH_SREAS2"), True)
        If Not IsNull(rsTB("SH_SREAS3")) Then rsWRK("TT_SALR3") = READTABLE("SDRC", rsTB("SH_SREAS3"), True)
        rsWRK("TT_SALPC2") = rsTB("SH_SALPC2")
        rsWRK("TT_SALPC3") = rsTB("SH_SALPC3")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 04, 07
    'FName.Caption = "HR_JOB_HISTORY"
    rsTB.Open "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR=" & xEmpNbr & _
            " ORDER BY JH_SDATE DESC,JH_CURRENT,JH_ID", gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then xDate0 = rsTB("JH_SDATE")
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        If rsTB("JH_CURRENT") Then
            rsWRK("TT_RECNBR") = "04"
        Else
            rsWRK("TT_RECNBR") = "07"
        End If
        rsWRK("TT_JOB") = ReadJob(rsTB("JH_JOB"))
        rsWRK("TT_GRID") = READTABLE("JBGD", rsTB("JH_GRID"), True)
        rsWRK("TT_JREAS") = READTABLE("SDRC", rsTB("JH_JREASON"), True)
        rsWRK("TT_REPTAU") = rsTB("JH_REPTAU")
        If xDate0 <> rsTB("JH_SDATE") Then
            rsWRK("TT_DATECHG") = xDate0
            'xNoDay# = DateDiff("d", RSTB("JH_SDATE"), xDate0)
            'rswrk("TT_NUMERIC") = xNoDay#
            xDate0 = rsTB("JH_SDATE")
        End If
        rsWRK("TT_SDATE") = rsTB("JH_SDATE")
        rsWRK("TT_SHIFT") = rsTB("JH_SHIFT")
        rsWRK("TT_FTE") = rsTB("JH_FTENUM")
        rsWRK("TT_FTEHRS") = rsTB("JH_FTEHRS")
        rsWRK("TT_WHRS") = rsTB("JH_WHRS")
        rsWRK("TT_DHRS") = rsTB("JH_DHRS")
        rsWRK("TT_PHRS") = rsTB("JH_PHRS")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 06, 09
    'FName.Caption = "HR_PERFORM_HISTORY"
    rsTB.Open "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR=" & xEmpNbr & _
            " ORDER BY PH_CURRENT,PH_PREVIEW DESC, PH_ID", gdbAdoIhr001, adOpenKeyset
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        If rsTB("PH_CURRENT") Then
            rsWRK("TT_RECNBR") = "06"
        Else
            rsWRK("TT_RECNBR") = "09"
        End If
        rsWRK("TT_JOB") = ReadJob(rsTB("PH_JOB"))
        rsWRK("TT_REPTAU") = rsTB("PH_REPTAU")
        If Not IsNull(rsTB("PH_PCODE")) Then rsWRK("TT_PCODE") = READTABLE("SDPC", rsTB("PH_PCODE"), True)
        rsWRK("TT_PREVIEW") = rsTB("PH_PREVIEW")
        rsWRK("TT_PNEXT") = rsTB("PH_PNEXT")
        'rswrk("TT_PCOMM") = RSTB("PH_COMMENTS")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 12
    'FName.Caption = "HREARN"
    rsTB.Open "SELECT * FROM HREARN WHERE EMPNBR=" & xEmpNbr & _
              " AND FDATE>=" & Date_SQL(xDate1) & _
              " AND TDATE<=" & Date_SQL(xDate2) & _
              " ORDER BY EARN_TYPE,TDATE DESC ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "Empl/Code"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "12"
            rsWRK("TT_COEFLAG") = 0
            rsWRK("TT_ETYPE") = "No Other Earnings"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "12"
            rsWRK("TT_ETYPE") = READTABLE("EARN", rsTB("EARN_TYPE"), True)
            rsWRK("TT_EFDATE") = rsTB("FDATE")
            rsWRK("TT_ETDATE") = rsTB("TDATE")
            rsWRK("TT_EACTUAL") = rsTB("ACT_DOLLAR")
            rsWRK("TT_COEFLAG") = rsTB("COST_OF_EMPLOYMENT")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 11
    'FName.Caption = "HREMPSKL"
    
    rsTB.Open "SELECT * FROM HREMPSKL WHERE SE_EMPNBR=" & xEmpNbr & " ORDER BY SE_DATE DESC", gdbAdoIhr001, adOpenKeyset
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "11"
            rsWRK("TT_SKILLD") = "No Skill"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "11"
            rsWRK("TT_SKILLD") = READTABLE("EDSK", rsTB("SE_SKILL"), True)
            rsWRK("TT_EXPFACT") = rsTB("SE_LEVEL")
            rsWRK("TT_SKLDTE") = rsTB("SE_DATE")
            '  rswrk("TT_SKLCOM") = RSTB("SE_COMM1")    'LAURA DEC 17, 1997
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 13
    'FName.Caption = "HREDU"
    rsTB.Open "SELECT * FROM HREDU WHERE EU_EMPNBR=" & xEmpNbr & _
              " ORDER BY EU_SCHOOL,EU_YEAR DESC ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL/CODE"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "13"
            rsWRK("TT_DEGREED") = "No Formal Education"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "13"
            rsWRK("TT_DEGREED") = READTABLE("EUDE", rsTB("EU_DEGREE"), False)
            rsWRK("TT_YEAR") = rsTB("EU_YEAR")
            rsWRK("TT_MAJORD") = READTABLE("EUMJ", rsTB("EU_MAJOR"), False)
            rsWRK("TT_MINORD") = READTABLE("EUMJ", rsTB("EU_MINOR"), False)
            rsWRK("TT_COMPL") = rsTB("EU_COMP")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 15

    rsTB.Open "SELECT * FROM HREDSEM WHERE ES_EMPNBR=" & xEmpNbr & _
            " ORDER BY ES_CTYPE,ES_DATCOMP DESC", gdbAdoIhr001, adOpenKeyset
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "15"
            rsWRK("TT_CTYPED") = "No Course/Seminar"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "15"
            rsWRK("TT_CTYPED") = READTABLE("ESCT", rsTB("ES_CTYPE"), False)
            rsWRK("TT_COURSE") = rsTB("ES_COURSE")
            rsWRK("TT_DATCOMP") = rsTB("ES_DATCOMP")
            rsWRK("TT_RESULTD") = READTABLE("ESRT", rsTB("ES_RESULTS"), False)
            rsWRK("TT_TBCO") = rsTB("ES_TBCO")
            rsWRK("TT_TBEMP") = rsTB("ES_TBEMP")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 14
    'FName.Caption = "HRTRADE"
    rsTB.Open "SELECT * FROM HRTRADE WHERE TD_EMPNBR=" & xEmpNbr & _
            " ORDER BY TD_CODE,TD_BEGINDT DESC", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL#/CODE"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "17"
            rsWRK("TT_CODED") = "No Association/Memberships"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "17"
            rsWRK("TT_CODED") = READTABLE("TDCD", rsTB("TD_CODE"), True)
            rsWRK("TT_DUES") = rsTB("TD_DUES")
            rsWRK("TT_COMPPD") = rsTB("TD_COMPPD")
            rsWRK("TT_RENEWDT") = rsTB("TD_RENEWDT")
            rsWRK("TT_BEGINDT") = rsTB("TD_BEGINDT")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 15
    'FName.Caption = "HRBENFT"
    rsTB.Open "SELECT * FROM HRBENFT WHERE BF_EMPNBR=" & xEmpNbr & _
            " ORDER BY BF_BCODE,BF_EDATE DESC,BF_BENE_ID", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL/CODE"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "19"
            rsWRK("TT_BCODED") = "No Benefit"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "19"
            rsWRK("TT_BCODED") = READTABLE("BNCD", rsTB("BF_BCODE"), True)
            rsWRK("TT_COVER") = rsTB("BF_COVER")
            rsWRK("TT_BEDATE") = rsTB("BF_EDATE")
            rsWRK("TT_BAMT") = rsTB("BF_AMT")
            rsWRK("TT_CCOST") = rsTB("BF_CCOST")
            rsWRK("TT_ECOST") = rsTB("BF_ECOST")
            rsWRK("TT_PCE") = rsTB("BF_PCE")
            rsWRK("TT_PCC") = rsTB("BF_PCC")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 21
    'FName.Caption = "HRBENS"
    rsTB.Open "SELECT * FROM HRBENS WHERE BD_EMPNBR=" & xEmpNbr & _
            " ORDER BY BD_BCODE,BD_ID", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "PrimaryKey"
    If rsTB.EOF Then 'Ticket #13685, if no record and then add a record to show it record
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "20"
            rsWRK("TT_BCODED") = "No Beneficiary"
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
    Else
        Do Until rsTB.EOF
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "20"
            rsWRK("TT_BCODED") = READTABLE("BNCD", rsTB("BD_BCODE"), True)
            rsWRK("TT_BNAME") = rsTB("BD_BNAME")
            rsWRK("TT_BRELATE") = rsTB("BD_RELATE")
            rsWRK("TT_BDOB") = rsTB("BD_DOB")
            rsWRK("TT_PCE") = rsTB("BD_PC")
            rsWRK("TT_WRKEMP") = glbUserID
            rsWRK.Update
            rsTB.MoveNext
            If rsTB.EOF Then Exit Do
        Loop
    End If
    rsTB.Close
    '--------------------------------------------------- 21
    xWCB = 0
    xOT = 0
    xCT = 0
    'FName.Caption = "HR_ATTENDANCE"
      
    rsTB.Open "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNbr & _
              " AND AD_DOA>=" & Date_SQL(xDate1) & _
              " AND AD_DOA<=" & Date_SQL(xDate2) & _
              " AND (AD_REASON like 'WCB%' OR AD_REASON like 'CT%' OR AD_REASON like 'OT%') " & _
              " ORDER BY AD_REASON,AD_DOA DESC", gdbAdoIhr001, adOpenKeyset
    
    'RSTB.Index = "EMPL/CODE"
    Do Until rsTB.EOF
        If Left(rsTB("AD_REASON"), 3) = "WCB" Then xWCB = xWCB + rsTB("AD_HRS")                                     '
        If Left(rsTB("AD_REASON"), 2) = "CT" Then xCT = xCT + rsTB("AD_HRS")                                        '
        If Left(rsTB("AD_REASON"), 2) = "OT" Then xOT = xOT + rsTB("AD_HRS")                                        '
        rsTB.MoveNext                                                                                                 '
        If rsTB.EOF Then Exit Do                                                                                      '
    Loop                                                                                                            '
    rsWRK.AddNew
    rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
    rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
    rsWRK("TT_RECNBR") = "21"
    rsWRK("TT_VAC") = rsEmp("ED_VAC")
    rsWRK("TT_SICK") = rsEmp("ED_SICK")
    rsWRK("TT_PVAC") = rsEmp("ED_PVAC")
    rsWRK("TT_PSICK") = rsEmp("ED_PSICK")
    rsWRK("TT_OTHRS") = xOT
    rsWRK("TT_CTHRS") = xCT
    rsWRK("TT_WCBHRS") = xWCB
    rsWRK("TT_VACHRS") = rsEmp("ED_VACT")
    rsWRK("TT_SICKHRS") = rsEmp("ED_SICKT")
    rsWRK("TT_WRKEMP") = glbUserID
    rsWRK.Update
    rsTB.Close
    '--------------------------------------------------22
    'FName.Caption = "HRENTHRS"
    rsTB.Open "SELECT * FROM HRENTHRS WHERE HE_EMPNBR=" & xEmpNbr & _
              " AND HE_FDATE>=" & Date_SQL(xDate1) & _
              " AND HE_FDATE<=" & Date_SQL(xDate2) & _
              " ORDER BY HE_TYPE,HE_TDATE DESC", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "PrimaryKey"
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        rsWRK("TT_RECNBR") = "22"
        rsWRK("TT_ETYPE") = READTABLE("ADRE", rsTB("HE_TYPE"), True)
        rsWRK("TT_EFDATE") = rsTB("HE_FDATE")
        rsWRK("TT_ETDATE") = rsTB("HE_TDATE")
        rsWRK("TT_ENTITLE") = rsTB("HE_ENTITLE")
        rsWRK("TT_EACTUAL") = rsTB("HE_TAKEN")
        rsWRK("TT_COEFLAG") = rsTB("HE_COE")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 23
    'FName.Caption = "HRDOLENT"
    rsTB.Open "SELECT * FROM HRDOLENT WHERE DE_EMPNBR=" & xEmpNbr & _
              " AND DE_FDATE>=" & Date_SQL(xDate1) & _
              " AND DE_FDATE<=" & Date_SQL(xDate2) & _
              " ORDER BY DE_TYPE,DE_TDATE DESC ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPL/CODE"
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        rsWRK("TT_RECNBR") = "23"
        rsWRK("TT_DTYPE") = READTABLE("EDOL", rsTB("DE_TYPE"), True)
        rsWRK("TT_DFDATE") = rsTB("DE_FDATE")
        rsWRK("TT_DTDATE") = rsTB("DE_TDATE")
        rsWRK("TT_DENTITL") = rsTB("DE_ENTITLE")
        rsWRK("TT_DACTUAL") = rsTB("DE_ACTUAL")
        rsWRK("TT_COEFLAG") = rsTB("DE_COST_OF_EMPLOYMENT")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 25
    'FName.Caption = "HREMPHIS"
    rsTB.Open "SELECT * FROM HREMPHIS WHERE EE_EMPNBR=" & xEmpNbr & _
              " ORDER BY EE_CHGDATE ", gdbAdoIhr001, adOpenKeyset
    'RSTB.Index = "EMPNBR"
    Do Until rsTB.EOF
        If Not IsNull(rsTB("EE_OLDDEPT")) Or Not IsNull(rsTB("EE_NEWDEPT")) Or Not IsNull(rsTB("EE_OLDDIV")) Or Not IsNull(rsTB("EE_NEWDIV")) Or Not IsNull(rsTB("EE_OLDSTAT")) Or Not IsNull(rsTB("EE_NEWSTAT")) Or Not IsNull(rsTB("EE_OLDORG")) Or Not IsNull(rsTB("EE_NEWORG")) Or Not IsNull(rsTB("EE_OLDPT")) Or Not IsNull(rsTB("EE_NEWPT")) Then
            rsWRK.AddNew
            rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
            rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
            rsWRK("TT_RECNBR") = "25"
            rsWRK("TT_OLDDEPT") = rsTB("EE_OLDDEPT")
            rsWRK("TT_NEWDEPT") = rsTB("EE_NEWDEPT")
            rsWRK("TT_OLDDIV") = rsTB("EE_OLDDIV")
            rsWRK("TT_NEWDIV") = rsTB("EE_NEWDIV")
            rsWRK("TT_OLDEMP") = rsTB("EE_OLDSTAT")
            rsWRK("TT_NEWEMP") = rsTB("EE_NEWSTAT")
            'rswrk("TT_OLDJOB") = RSTB("EE_OLDJOB")
            'rswrk("TT_NEWJOB") = RSTB("EE_NEWJOB")
            rsWRK("TT_OLDPT") = rsTB("EE_OLDPT")
            rsWRK("TT_NEWPT") = rsTB("EE_NEWPT")
            rsWRK("TT_OLDORG") = rsTB("EE_OLDORG")
            rsWRK("TT_NEWORG") = rsTB("EE_NEWORG")
            'rswrk("TT_OLDGL") = RSTB("EE_OLDGLNO")
            'rswrk("TT_NEWGL") = RSTB("EE_NEWGLNO")
            'rswrk("TT_FLAG") = RSTB("EE_DOTFLAG")
            rsWRK("TT_CHGDATE") = rsTB("EE_CHGDATE")
            rsWRK("TT_WRKEMP") = glbUserID
            'Hemu
            rsWRK("TT_TODATE") = rsTB("EE_TODATE")
            'Hemu
            rsWRK.Update
        End If
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '--------------------------------------------------- 26
    'FName.Caption = "HRSTATUS"
    rsTB.Open "SELECT * FROM HRSTATUS WHERE SC_EMPNBR=" & xEmpNbr & _
              " ORDER BY SC_LDATE,SC_LTIME ", gdbAdoIhr001, adOpenKeyset
    Do Until rsTB.EOF
        rsWRK.AddNew
        rsWRK("TT_COMPNO") = rsEmp("ED_COMPNO")
        rsWRK("TT_EMPNBR") = rsEmp("ED_EMPNBR")
        rsWRK("TT_RECNBR") = "26"
        rsWRK("TT_OLDEMP") = rsTB("SC_OLDEMP")
        rsWRK("TT_NEWEMP") = rsTB("SC_NEWEMP")
        rsWRK("TT_EFDATE") = rsTB("SC_FDATE")
        rsWRK("TT_ETDATE") = rsTB("SC_TDATE")
        rsWRK("TT_JOBCODE") = rsTB("SC_JOB")
        rsWRK("TT_EMP") = rsTB("SC_REASON")
        rsWRK("TT_WRKEMP") = glbUserID
        rsWRK.Update
        rsTB.MoveNext
        If rsTB.EOF Then Exit Do
    Loop
    rsTB.Close
    '------------------------------------------------ end
    rsEmp.MoveNext
    If rsEmp.EOF Then Exit Do
Loop
rsEmp.Close
rsWRK.Close
Exit Sub

ERR_EmpWrkOracle:
If Err = 13 Then
  MsgBox "SYSTEM ERROR : 13 - Type MisMatch"
Else
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
End If
If False Then
    Resume  ' for debugging
End If
End Sub

Public Sub ReCalcOvt(WSQLQ)
Dim rsTA As New ADODB.Recordset
Dim SQLQ
Dim rsOvtMst As New ADODB.Recordset
Dim rsHREmp As New ADODB.Recordset
Dim rsOvtBank As New ADODB.Recordset
On Error GoTo ReCalcOvt_Err

Screen.MousePointer = vbHourglass

MDIMain.panHelp(0).FloodType = 1            '28July99 js
MDIMain.panHelp(1).Caption = " Please Wait" '
MDIMain.panHelp(2).Caption = ""             '
MDIMain.panHelp(0).FloodPercent = 10

SQLQ = "Update HR_OVERTIME_BANK SET OT_BANK=0, OT_BANKT=0"
If Len(Trim(WSQLQ)) <> 0 Then
    SQLQ = SQLQ & " WHERE " & WSQLQ
End If
gdbAdoIhr001.Execute SQLQ

If glbOracle Then
    'Overtime Banked - Current
    SQLQ = " Update HR_OVERTIME_BANK SET "
    SQLQ = SQLQ & " HR_OVERTIME_BANK.OT_BANK =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " Where OT_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND (AD_DOA>= OT_EFDATE) AND (AD_DOA<=OT_ETDATE)"
    SQLQ = SQLQ & " AND (AD_REASON Like 'OT%') AND (AD_REASON <> 'OTBF') )" 'Ticket 16064
    SQLQ = SQLQ & " WHERE OT_EMPNBR IN"
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    SQLQ = SQLQ & " AND (AD_DOA>= OT_EFDATE) AND (AD_DOA<=OT_ETDATE)"
    SQLQ = SQLQ & " AND (AD_REASON Like 'OT%') AND (AD_REASON <> 'OTBF') )" 'Ticket 16064
    
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 15
    
    
    'Overtime Banked - Previous
    SQLQ = " Update HR_OVERTIME_BANK SET "
    SQLQ = SQLQ & " HR_OVERTIME_BANK.OT_PBANK =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " Where OT_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND (AD_DOA>= OT_EFDATE) AND (AD_DOA<=OT_ETDATE)"
    SQLQ = SQLQ & " AND (AD_REASON = 'OTBF') )" 'Ticket 16064
    SQLQ = SQLQ & " WHERE OT_EMPNBR IN"
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    SQLQ = SQLQ & " AND (AD_DOA=OT_EFDATE) "
    SQLQ = SQLQ & " AND (AD_REASON = 'OTBF') )" 'Ticket 16064
    
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 30
    
    'Update with Multiplier
    SQLQ = "SELECT * FROM HR_OVERTIME_MASTER"
    rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsOvtMst.EOF Then
        rsOvtMst.MoveFirst
        Do While Not rsOvtMst.EOF
            SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 "
            If Not IsNull(rsOvtMst("OM_ORG")) And rsOvtMst("OM_ORG") <> "" Then
                SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_PT")) And rsOvtMst("OM_PT") <> "" Then
                SQLQ = SQLQ & " AND ED_PT = '" & rsOvtMst("OM_PT") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_EMP")) And rsOvtMst("OM_EMP") <> "" Then
                SQLQ = SQLQ & " AND ED_EMP = '" & rsOvtMst("OM_EMP") & "'"
            End If
                        
            'Ticket #15753
            If Not IsNull(rsOvtMst("OM_LOC")) And rsOvtMst("OM_LOC") <> "" Then
                SQLQ = SQLQ & " AND ED_LOC = '" & rsOvtMst("OM_LOC") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_REGION")) And rsOvtMst("OM_REGION") <> "" Then
                SQLQ = SQLQ & " AND ED_REGION = '" & rsOvtMst("OM_REGION") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_ADMINBY")) And rsOvtMst("OM_ADMINBY") <> "" Then
                SQLQ = SQLQ & " AND ED_ADMINBY = '" & rsOvtMst("OM_ADMINBY") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_SECTION")) And rsOvtMst("OM_SECTION") <> "" Then
                SQLQ = SQLQ & " AND ED_SECTION = '" & rsOvtMst("OM_SECTION") & "'"
            End If

            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHREmp.EOF Then
                Do While Not rsHREmp.EOF
                    SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & rsHREmp("ED_EMPNBR")
                    
                    If Len(Trim(WSQLQ)) <> 0 Then
                        SQLQ = SQLQ & " AND " & WSQLQ
                    End If
                    
                    rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsOvtBank.EOF Then
                        rsOvtBank("OT_BANK") = rsOvtBank("OT_BANK") * rsOvtMst("OM_MULTIPLIER")
                        rsOvtBank.Update
                    End If
                    rsOvtBank.Close
                    
                    rsHREmp.MoveNext
                Loop
            End If
            rsHREmp.Close
            
            rsOvtMst.MoveNext
        Loop
    End If
    rsOvtMst.Close
    
    'SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_BANK = OT_BANK * "
    'SQLQ = SQLQ & " (SELECT OM_MULTIPLIER FROM HR_OVERTIME_MASTER WHERE OM_ORG ="
    'SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = OT_EMPNBR))"
    
    'If Len(Trim(WSQLQ)) <> 0 Then
    '    SQLQ = SQLQ & " WHERE " & WSQLQ
    'End If
    'gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 50
    
    'Overtime Taken
    SQLQ = " Update HR_OVERTIME_BANK SET "
    SQLQ = SQLQ & " HR_OVERTIME_BANK.OT_BANKT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " Where OT_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND (AD_DOA>= OT_EFDATE) AND (AD_DOA<=OT_ETDATE)"
    SQLQ = SQLQ & " AND (AD_REASON Like 'CT%') )"
    SQLQ = SQLQ & " WHERE OT_EMPNBR IN"
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE WHERE HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    SQLQ = SQLQ & " AND (AD_DOA>= OT_EFDATE) AND (AD_DOA<=OT_ETDATE)"
    SQLQ = SQLQ & " AND (AD_REASON Like 'CT%') )"
    
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 70
    
    'Update with Maximum if Blank or Null
    'Modified by Sam as it was not updating records other than Blank or Null 07/12/2006
    'Also for single employee it should not update MaxBank value from master table
            
        
    If Len(Trim(WSQLQ)) = 0 Then
        SQLQ = "SELECT * FROM HR_OVERTIME_MASTER"
        rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsOvtMst.EOF Then
            rsOvtMst.MoveFirst
            Do While Not rsOvtMst.EOF
                SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 "
                If Not IsNull(rsOvtMst("OM_ORG")) And rsOvtMst("OM_ORG") <> "" Then
                    SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_PT")) And rsOvtMst("OM_PT") <> "" Then
                    SQLQ = SQLQ & " AND ED_PT = '" & rsOvtMst("OM_PT") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_EMP")) And rsOvtMst("OM_EMP") <> "" Then
                    SQLQ = SQLQ & " AND ED_EMP = '" & rsOvtMst("OM_EMP") & "'"
                End If
                            
                'Ticket #15753
                If Not IsNull(rsOvtMst("OM_LOC")) And rsOvtMst("OM_LOC") <> "" Then
                    SQLQ = SQLQ & " AND ED_LOC = '" & rsOvtMst("OM_LOC") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_REGION")) And rsOvtMst("OM_REGION") <> "" Then
                    SQLQ = SQLQ & " AND ED_REGION = '" & rsOvtMst("OM_REGION") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_ADMINBY")) And rsOvtMst("OM_ADMINBY") <> "" Then
                    SQLQ = SQLQ & " AND ED_ADMINBY = '" & rsOvtMst("OM_ADMINBY") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_SECTION")) And rsOvtMst("OM_SECTION") <> "" Then
                    SQLQ = SQLQ & " AND ED_SECTION = '" & rsOvtMst("OM_SECTION") & "'"
                End If
                            
                rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHREmp.EOF Then
                    Do While Not rsHREmp.EOF
                        SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & rsHREmp("ED_EMPNBR")
                        rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsOvtBank.EOF Then
                            rsOvtBank("OT_MBANK") = rsOvtMst("OM_MAX_BANK_HRS")
                            rsOvtBank.Update
                        End If
                        rsOvtBank.Close
                        
                        rsHREmp.MoveNext
                    Loop
                End If
                rsHREmp.Close
                
                rsOvtMst.MoveNext
            Loop
        End If
        rsOvtMst.Close
    
        'SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_MBANK = "
        'SQLQ = SQLQ & " (SELECT OM_MAX_BANK_HRS FROM HR_OVERTIME_MASTER WHERE OM_ORG ="
        'SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = OT_EMPNBR))" 'WHERE (OT_MBANK IS NULL) "
    End If
    'gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 100
    
ElseIf glbSQL Then
    'Overtime Bank - Current
    SQLQ = " Update HR_OVERTIME_BANK SET "
    SQLQ = SQLQ & " HR_OVERTIME_BANK.OT_BANK =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " WHERE OT_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND AD_DOA BETWEEN OT_EFDATE AND OT_ETDATE"
    SQLQ = SQLQ & " AND AD_REASON Like 'OT%' AND AD_REASON <> 'OTBF')"  'Ticket 16064
    SQLQ = SQLQ & " WHERE OT_EMPNBR IN"
   'Commented by Sam as it was using wrong table HREMP instead of HR_OVERTIME_BANK
   ' SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HR_OVERTIME_BANK ON HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN OT_EFDATE AND OT_ETDATE)"
    SQLQ = SQLQ & " AND AD_REASON Like 'OT%' AND AD_REASON <> 'OTBF')"  'Ticket 16064
    
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 15
    
    'Overtime Bank - Previous
    SQLQ = " Update HR_OVERTIME_BANK SET "
    SQLQ = SQLQ & " HR_OVERTIME_BANK.OT_PBANK =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " WHERE OT_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND AD_DOA BETWEEN OT_EFDATE AND OT_ETDATE"
    SQLQ = SQLQ & " AND AD_REASON = 'OTBF')"  'Ticket 16064
    SQLQ = SQLQ & " WHERE OT_EMPNBR IN"
   'Commented by Sam as it was using wrong table HREMP instead of HR_OVERTIME_BANK
   ' SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HR_OVERTIME_BANK ON HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    SQLQ = SQLQ & " WHERE (AD_DOA = OT_EFDATE )"
    SQLQ = SQLQ & " AND AD_REASON = 'OTBF')"  'Ticket 16064
    
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 30
    
    'Update with Multiplier
    SQLQ = "SELECT * FROM HR_OVERTIME_MASTER"
    rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsOvtMst.EOF Then
        rsOvtMst.MoveFirst
        Do While Not rsOvtMst.EOF
            SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 "
            If Not IsNull(rsOvtMst("OM_ORG")) And rsOvtMst("OM_ORG") <> "" Then
                SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_PT")) And rsOvtMst("OM_PT") <> "" Then
                SQLQ = SQLQ & " AND ED_PT = '" & rsOvtMst("OM_PT") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_EMP")) And rsOvtMst("OM_EMP") <> "" Then
                SQLQ = SQLQ & " AND ED_EMP = '" & rsOvtMst("OM_EMP") & "'"
            End If
                        
            'Ticket #15753
            If Not IsNull(rsOvtMst("OM_LOC")) And rsOvtMst("OM_LOC") <> "" Then
                SQLQ = SQLQ & " AND ED_LOC = '" & rsOvtMst("OM_LOC") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_REGION")) And rsOvtMst("OM_REGION") <> "" Then
                SQLQ = SQLQ & " AND ED_REGION = '" & rsOvtMst("OM_REGION") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_ADMINBY")) And rsOvtMst("OM_ADMINBY") <> "" Then
                SQLQ = SQLQ & " AND ED_ADMINBY = '" & rsOvtMst("OM_ADMINBY") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_SECTION")) And rsOvtMst("OM_SECTION") <> "" Then
                SQLQ = SQLQ & " AND ED_SECTION = '" & rsOvtMst("OM_SECTION") & "'"
            End If
                        
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHREmp.EOF Then
                Do While Not rsHREmp.EOF
                    SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & rsHREmp("ED_EMPNBR")
                    
                    If Len(Trim(WSQLQ)) <> 0 Then
                        SQLQ = SQLQ & " AND " & WSQLQ
                    End If
                    
                    rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsOvtBank.EOF Then
                        rsOvtBank("OT_BANK") = rsOvtBank("OT_BANK") * rsOvtMst("OM_MULTIPLIER")
                        rsOvtBank.Update
                    End If
                    rsOvtBank.Close
                    
                    rsHREmp.MoveNext
                Loop
            End If
            rsHREmp.Close
            
            rsOvtMst.MoveNext
        Loop
    End If
    rsOvtMst.Close
    
    'SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_BANK = OT_BANK * "
    'SQLQ = SQLQ & " (SELECT OM_MULTIPLIER FROM HR_OVERTIME_MASTER WHERE OM_ORG ="
    'SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = OT_EMPNBR))"
    
    'If Len(Trim(WSQLQ)) <> 0 Then
    '    SQLQ = SQLQ & " WHERE " & WSQLQ
    'End If
    'gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 50
        
        
    'Overtime Taken
    SQLQ = " Update HR_OVERTIME_BANK SET "
    SQLQ = SQLQ & " HR_OVERTIME_BANK.OT_BANKT =(SELECT SUM(AD_HRS) FROM HR_ATTENDANCE"
    SQLQ = SQLQ & " WHERE OT_EMPNBR = AD_EMPNBR"
    SQLQ = SQLQ & " AND AD_DOA BETWEEN OT_EFDATE AND OT_ETDATE"
    SQLQ = SQLQ & " AND AD_REASON Like 'CT%')"
    SQLQ = SQLQ & " WHERE OT_EMPNBR IN"
    'Commented by Sam as it was using wrong table HREMP instead of HR_OVERTIME_BANK
   ' SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HREMP ON HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    SQLQ = SQLQ & " (SELECT AD_EMPNBR FROM HR_ATTENDANCE INNER JOIN HR_OVERTIME_BANK ON HR_ATTENDANCE.AD_EMPNBR=HR_OVERTIME_BANK.OT_EMPNBR"
    SQLQ = SQLQ & " WHERE (AD_DOA BETWEEN OT_EFDATE AND OT_ETDATE)"
    SQLQ = SQLQ & " AND AD_REASON Like 'CT%')"
    
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    
    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 70
    
    'Update with Maximum if Blank or Null
    'Modified by Sam as it was not updating records other than Blank or Null 07/12/2006
    'Also for single employee it should not update MaxBank value from master table
    If Len(Trim(WSQLQ)) = 0 Then
        SQLQ = "SELECT * FROM HR_OVERTIME_MASTER"
        rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsOvtMst.EOF Then
            rsOvtMst.MoveFirst
            Do While Not rsOvtMst.EOF
                SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 "
                If Not IsNull(rsOvtMst("OM_ORG")) And rsOvtMst("OM_ORG") <> "" Then
                    SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_PT")) And rsOvtMst("OM_PT") <> "" Then
                    SQLQ = SQLQ & " AND ED_PT = '" & rsOvtMst("OM_PT") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_EMP")) And rsOvtMst("OM_EMP") <> "" Then
                    SQLQ = SQLQ & " AND ED_EMP = '" & rsOvtMst("OM_EMP") & "'"
                End If
                            
                'Ticket #15753
                If Not IsNull(rsOvtMst("OM_LOC")) And rsOvtMst("OM_LOC") <> "" Then
                    SQLQ = SQLQ & " AND ED_LOC = '" & rsOvtMst("OM_LOC") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_REGION")) And rsOvtMst("OM_REGION") <> "" Then
                    SQLQ = SQLQ & " AND ED_REGION = '" & rsOvtMst("OM_REGION") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_ADMINBY")) And rsOvtMst("OM_ADMINBY") <> "" Then
                    SQLQ = SQLQ & " AND ED_ADMINBY = '" & rsOvtMst("OM_ADMINBY") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_SECTION")) And rsOvtMst("OM_SECTION") <> "" Then
                    SQLQ = SQLQ & " AND ED_SECTION = '" & rsOvtMst("OM_SECTION") & "'"
                End If
                            
                rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHREmp.EOF Then
                    Do While Not rsHREmp.EOF
                        SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & rsHREmp("ED_EMPNBR")
                        rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsOvtBank.EOF Then
                            rsOvtBank("OT_MBANK") = rsOvtMst("OM_MAX_BANK_HRS")
                            rsOvtBank.Update
                        End If
                        rsOvtBank.Close
                        
                        rsHREmp.MoveNext
                    Loop
                End If
                rsHREmp.Close
                
                rsOvtMst.MoveNext
            Loop
        End If
        rsOvtMst.Close
   
        'SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_MBANK = "
        'SQLQ = SQLQ & " (SELECT OM_MAX_BANK_HRS FROM HR_OVERTIME_MASTER WHERE OM_ORG ="
        'SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = OT_EMPNBR))" 'WHERE (OT_MBANK IS NULL) "
    End If
    'gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 100
    
Else
    'Overtime Banked - Current
    'Commented by Sam as it was not working in Access 07/11/2006
    'SQLQ = "SELECT OT_EMPNBR, Sum(AD_HRS) AS SumHRS"
    'SQLQ = SQLQ & " FROM HR_OVERTIME_BANK INNER JOIN HR_ATTENDANCE ON HR_OVERTIME_BANK.OT_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
    'SQLQ = SQLQ & " WHERE AD_DOA>=OT_EFDATE And AD_DOA<=OT_ETDATE AND LEFT(AD_REASON,3)='OT' "
'    If Len(Trim(WSQLQ)) <> 0 Then
'        SQLQ = SQLQ & " AND " & WSQLQ
'    End If
    'COMMENTED BY SAM 07/11/2006
    'SQLQ = SQLQ & " GROUP BY OT_EMPNBR "
    'REDONE BY SAM AS IT WAS NOT WORKING FOR ACCESS 07/11/2006
    'SQLQ = "SELECT OT_EMPNBR, Sum(AD_HRS) AS SumHRS, AD_DOA,OT_EFDATE,OT_ETDATE"
    'SQLQ = SQLQ & " FROM HR_OVERTIME_BANK INNER JOIN HR_ATTENDANCE ON HR_OVERTIME_BANK.OT_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
    'SQLQ = SQLQ & " GROUP BY AD_REASON,OT_EMPNBR,AD_DOA,OT_EFDATE,OT_ETDATE"
    'SQLQ = SQLQ & " HAVING YEAR(AD_DOA) = " & Year(Date)
    'SQLQ = SQLQ & " AND LEFT(AD_REASON,2)='OT'"
    
    'The above query does not work either - Hemu
    'It's not calculating the Sum correctly because of Group By which does not allow Sum of OTs
    SQLQ = "SELECT OT_EMPNBR, Sum(AD_HRS) AS SumHRS "
    SQLQ = SQLQ & " FROM HR_OVERTIME_BANK INNER JOIN HR_ATTENDANCE ON HR_OVERTIME_BANK.OT_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
    'SQLQ = SQLQ & " WHERE YEAR(AD_DOA) = " & Year(Date)
    SQLQ = SQLQ & " WHERE (AD_DOA >= OT_EFDATE AND AD_DOA <= OT_ETDATE)"
    SQLQ = SQLQ & " AND LEFT(AD_REASON,2)='OT' AND (AD_REASON <> 'OTBF')"   'Ticket 16064
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    SQLQ = SQLQ & " GROUP BY OT_EMPNBR"
        
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenDynamic
    If Not rsTA.EOF Then
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HR_OVERTIME_BANK SET OT_BANK=" & rsTA("SUMHRS") & " WHERE OT_EMPNBR=" & rsTA("OT_EMPNBR")
            rsTA.MoveNext
        Loop
    End If
    rsTA.Close
    MDIMain.panHelp(0).FloodPercent = 15
    
    'Overtime Banked - Previous
    SQLQ = "SELECT OT_EMPNBR, Sum(AD_HRS) AS SumHRS "
    SQLQ = SQLQ & " FROM HR_OVERTIME_BANK INNER JOIN HR_ATTENDANCE ON HR_OVERTIME_BANK.OT_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
    'SQLQ = SQLQ & " WHERE YEAR(AD_DOA) = " & Year(Date)
    SQLQ = SQLQ & " WHERE (AD_DOA >= OT_EFDATE AND AD_DOA <= OT_ETDATE)"
    SQLQ = SQLQ & " AND (AD_REASON = 'OTBF')"   'Ticket 16064
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    SQLQ = SQLQ & " GROUP BY OT_EMPNBR"
        
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenDynamic
    If Not rsTA.EOF Then
        Do Until rsTA.EOF
            gdbAdoIhr001.Execute "UPDATE HR_OVERTIME_BANK SET OT_PBANK=" & rsTA("SUMHRS") & " WHERE OT_EMPNBR=" & rsTA("OT_EMPNBR")
            rsTA.MoveNext
        Loop
    End If
    rsTA.Close
    MDIMain.panHelp(0).FloodPercent = 30
    
    
    'Update with Multiplier
    'REDONE BY SAM AS IT WAS NOT WORKING FOR ACCESS 07/11/2006
'    SQLQ = "SELECT HR_OVERTIME_MASTER.OM_MULTIPLIER, HR_OVERTIME_BANK.OT_EMPNBR"
'    SQLQ = SQLQ & " FROM HR_OVERTIME_MASTER INNER JOIN "
'    SQLQ = SQLQ & "(HREMP INNER JOIN HR_OVERTIME_BANK "
'    SQLQ = SQLQ & "ON HREMP.ED_EMPNBR = HR_OVERTIME_BANK.OT_EMPNBR) "
'    SQLQ = SQLQ & "ON (HR_OVERTIME_MASTER.OM_ORG_TABL = HREMP.ED_ORG_TABL) "
'    SQLQ = SQLQ & "AND (HR_OVERTIME_MASTER.OM_ORG = HREMP.ED_ORG)"
'    If Len(Trim(WSQLQ)) <> 0 Then
'        SQLQ = SQLQ & " WHERE " & WSQLQ
'    End If
'
'    rsta.Open SQLQ, gdbAdoIhr001, adOpenDynamic
'    If Not rsta.EOF Then
'        Do Until rsta.EOF
'        SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_BANK= OT_BANK *" & rsta("OM_MULTIPLIER")
'       If Len(Trim(WSQLQ)) <> 0 Then
'            SQLQ = SQLQ & " WHERE " & WSQLQ
'        Else
'            SQLQ = SQLQ & " WHERE OT_EMPNBR =" & rsta("OT_EMPNBR")
'        End If
'
'            gdbAdoIhr001.Execute SQLQ
'            rsta.MoveNext
'        Loop
'    End If
'    rsta.Close
        
    SQLQ = "SELECT * FROM HR_OVERTIME_MASTER"
    rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsOvtMst.EOF Then
        rsOvtMst.MoveFirst
        Do While Not rsOvtMst.EOF
            SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 "
            If Not IsNull(rsOvtMst("OM_ORG")) And rsOvtMst("OM_ORG") <> "" Then
                SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_PT")) And rsOvtMst("OM_PT") <> "" Then
                SQLQ = SQLQ & " AND ED_PT = '" & rsOvtMst("OM_PT") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_EMP")) And rsOvtMst("OM_EMP") <> "" Then
                SQLQ = SQLQ & " AND ED_EMP = '" & rsOvtMst("OM_EMP") & "'"
            End If
            
            'Ticket #15753
            If Not IsNull(rsOvtMst("OM_LOC")) And rsOvtMst("OM_LOC") <> "" Then
                SQLQ = SQLQ & " AND ED_LOC = '" & rsOvtMst("OM_LOC") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_REGION")) And rsOvtMst("OM_REGION") <> "" Then
                SQLQ = SQLQ & " AND ED_REGION = '" & rsOvtMst("OM_REGION") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_ADMINBY")) And rsOvtMst("OM_ADMINBY") <> "" Then
                SQLQ = SQLQ & " AND ED_ADMINBY = '" & rsOvtMst("OM_ADMINBY") & "'"
            End If
            If Not IsNull(rsOvtMst("OM_SECTION")) And rsOvtMst("OM_SECTION") <> "" Then
                SQLQ = SQLQ & " AND ED_SECTION = '" & rsOvtMst("OM_SECTION") & "'"
            End If
            
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsHREmp.EOF Then
                Do While Not rsHREmp.EOF
                    SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & rsHREmp("ED_EMPNBR")
                    
                    If Len(Trim(WSQLQ)) <> 0 Then
                        SQLQ = SQLQ & " AND " & WSQLQ
                    End If
                    
                    rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsOvtBank.EOF Then
                        rsOvtBank("OT_BANK") = rsOvtBank("OT_BANK") * rsOvtMst("OM_MULTIPLIER")
                        rsOvtBank.Update
                    End If
                    rsOvtBank.Close
                    
                    rsHREmp.MoveNext
                Loop
            End If
            rsHREmp.Close
            
            rsOvtMst.MoveNext
        Loop
    End If
    rsOvtMst.Close
    MDIMain.panHelp(0).FloodPercent = 50
    
    'Overtime Taken
    SQLQ = "SELECT OT_EMPNBR, Sum(AD_HRS) AS SumHRS"
    SQLQ = SQLQ & " FROM HR_OVERTIME_BANK INNER JOIN HR_ATTENDANCE ON HR_OVERTIME_BANK.OT_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
    'Ticket #21472 - to fix the error - moved from HAVING clause below
    SQLQ = SQLQ & " WHERE (AD_DOA >= OT_EFDATE AND AD_DOA <= OT_ETDATE)"
    SQLQ = SQLQ & " AND LEFT(AD_REASON,2)='CT'"
    
    'Ticket #21472 - to fix the error - moved from below
    If Len(Trim(WSQLQ)) <> 0 Then
        SQLQ = SQLQ & " AND " & WSQLQ
    End If
    
    SQLQ = SQLQ & " GROUP BY Year((AD_DOA)),AD_REASON,OT_EMPNBR"
    'SQLQ = SQLQ & " HAVING YEAR(AD_DOA) = " & Year(Date)
    'Ticket #21472 - to fix the error - moved to WHERE clause above
    'SQLQ = SQLQ & " HAVING (AD_DOA >= OT_EFDATE AND AD_DOA <= OT_ETDATE)"
    'SQLQ = SQLQ & " AND LEFT(AD_REASON,2)='CT'"
    
    'COMMENTED BY SAM AS IT WAS GIVING ERRORS
'    SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_BANK = OT_BANK * "
'
'    SQLQ = SQLQ & " (SELECT OM_MULTIPLIER FROM HR_OVERTIME_MASTER WHERE OM_ORG ="
'   SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = OT_EMPNBR))"
'
'    If Len(Trim(WSQLQ)) <> 0 Then
'        SQLQ = SQLQ & " WHERE " & WSQLQ
'    End If
'    gdbAdoIhr001.Execute SQLQ
'    MDIMain.panHelp(0).FloodPercent = 50

      
    'Overtime Taken
'    SQLQ = "SELECT OT_EMPNBR, Sum(AD_HRS) AS SumHRS"
'    SQLQ = SQLQ & " FROM HR_OVERTIME_BANK INNER JOIN HR_ATTENDANCE ON HR_OVERTIME_BANK.OT_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
'    SQLQ = SQLQ & " WHERE AD_DOA>=OT_EFDATE And AD_DOA<=OT_ETDATE AND LEFT(AD_REASON,3)='CT' "
    
    'Ticket #21472 - to fix the error - moved to above
    'If Len(Trim(WSQLQ)) <> 0 Then
    '    SQLQ = SQLQ & " AND " & WSQLQ
    'End If

    'SQLQ = SQLQ & " GROUP BY OT_EMPNBR "
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenDynamic
    Do Until rsTA.EOF
        gdbAdoIhr001.Execute "UPDATE HR_OVERTIME_BANK SET OT_BANKT=" & rsTA("SUMHRS") & " WHERE OT_EMPNBR=" & rsTA("OT_EMPNBR")
        rsTA.MoveNext
    Loop
    rsTA.Close
    MDIMain.panHelp(0).FloodPercent = 70
    
    'Update with Maximum for Single or All Employees
    'If it's Single then It will update Max Overtime from Overtime bank table Otherwise from Master table
    If Len(Trim(WSQLQ)) = 0 Then
        SQLQ = "SELECT * FROM HR_OVERTIME_MASTER"
        rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsOvtMst.EOF Then
            rsOvtMst.MoveFirst
            Do While Not rsOvtMst.EOF
                SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE 1 = 1 "
                If Not IsNull(rsOvtMst("OM_ORG")) And rsOvtMst("OM_ORG") <> "" Then
                    SQLQ = SQLQ & " AND ED_ORG = '" & rsOvtMst("OM_ORG") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_PT")) And rsOvtMst("OM_PT") <> "" Then
                    SQLQ = SQLQ & " AND ED_PT = '" & rsOvtMst("OM_PT") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_EMP")) And rsOvtMst("OM_EMP") <> "" Then
                    SQLQ = SQLQ & " AND ED_EMP = '" & rsOvtMst("OM_EMP") & "'"
                End If
                
                'Ticket #15753
                If Not IsNull(rsOvtMst("OM_LOC")) And rsOvtMst("OM_LOC") <> "" Then
                    SQLQ = SQLQ & " AND ED_LOC = '" & rsOvtMst("OM_LOC") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_REGION")) And rsOvtMst("OM_REGION") <> "" Then
                    SQLQ = SQLQ & " AND ED_REGION = '" & rsOvtMst("OM_REGION") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_ADMINBY")) And rsOvtMst("OM_ADMINBY") <> "" Then
                    SQLQ = SQLQ & " AND ED_ADMINBY = '" & rsOvtMst("OM_ADMINBY") & "'"
                End If
                If Not IsNull(rsOvtMst("OM_SECTION")) And rsOvtMst("OM_SECTION") <> "" Then
                    SQLQ = SQLQ & " AND ED_SECTION = '" & rsOvtMst("OM_SECTION") & "'"
                End If
                
                rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsHREmp.EOF Then
                    Do While Not rsHREmp.EOF
                        SQLQ = "SELECT * FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & rsHREmp("ED_EMPNBR")
                        rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If Not rsOvtBank.EOF Then
                            rsOvtBank("OT_MBANK") = rsOvtMst("OM_MAX_BANK_HRS")
                            rsOvtBank.Update
                        End If
                        rsOvtBank.Close
                        
                        rsHREmp.MoveNext
                    Loop
                End If
                rsHREmp.Close
                
                rsOvtMst.MoveNext
            Loop
        End If
        rsOvtMst.Close
    End If
    
'    If Len(Trim(WSQLQ)) <> 0 Then
'        SQLQ = "SELECT OT_MBANK"
'        SQLQ = SQLQ & " FROM HR_OVERTIME_BANK"
'        SQLQ = SQLQ & " WHERE " & WSQLQ
'
'    Else
'        SQLQ = "SELECT HR_OVERTIME_MASTER.OM_MAX_BANK_HRS, HR_OVERTIME_BANK.OT_EMPNBR"
'        SQLQ = SQLQ & " FROM HR_OVERTIME_MASTER INNER JOIN "
'        SQLQ = SQLQ & "(HREMP INNER JOIN HR_OVERTIME_BANK "
'        SQLQ = SQLQ & "ON HREMP.ED_EMPNBR = HR_OVERTIME_BANK.OT_EMPNBR) "
'        SQLQ = SQLQ & "ON (HR_OVERTIME_MASTER.OM_ORG_TABL = HREMP.ED_ORG_TABL) "
'        SQLQ = SQLQ & "AND (HR_OVERTIME_MASTER.OM_ORG = HREMP.ED_ORG)"
'
'    End If
'    rsta.Open SQLQ, gdbAdoIhr001, adOpenDynamic
'    If Not rsta.EOF Then
'        Do Until rsta.EOF
'            If Len(Trim(WSQLQ)) <> 0 Then
'               SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_MBANK=" & rsta("OT_MBANK")
'               SQLQ = SQLQ & " WHERE " & WSQLQ
'            Else
'                SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_MBANK=" & rsta("OM_MAX_BANK_HRS")
'                SQLQ = SQLQ & " WHERE OT_EMPNBR =" & rsta("OT_EMPNBR")
'            End If
'
'             gdbAdoIhr001.Execute SQLQ
'             rsta.MoveNext
'       Loop
'    End If
'    rsta.Close
    
 'COMMENTED BY SAM BECAUSE IT WAS NOT WORKING
'    SQLQ = "UPDATE HR_OVERTIME_BANK SET OT_MBANK = "
'    SQLQ = SQLQ & " (SELECT OM_MAX_BANK_HRS FROM HR_OVERTIME_MASTER WHERE OM_ORG ="
'    SQLQ = SQLQ & " (SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR = OT_EMPNBR)) WHERE (OT_MBANK IS NULL) "
    
'    If Len(Trim(WSQLQ)) <> 0 Then
'        SQLQ = SQLQ & " AND " & WSQLQ
'    End If
'    gdbAdoIhr001.Execute SQLQ
    MDIMain.panHelp(0).FloodPercent = 100
    
End If

'added by Bryan Ticket#11392 to recalculate on attendance screen
SQLQ = "SELECT OT_BANK, OT_BANKT, OT_EMPNBR FROM HR_OVERTIME_BANK"
If Len(Trim(WSQLQ)) <> 0 Then
    SQLQ = SQLQ & " WHERE " & WSQLQ
End If
rsTA.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
If Not rsTA.EOF Then
    Do Until rsTA.EOF
        If Not IsNull(rsTA("OT_BANK")) Then
            'Ticket #23655 - For City of Niagara Falls - ED_OTBANK is the total Banked Time.
            If glbCompSerial = "S/N - 2276W" Then
                SQLQ = "UPDATE HREMP SET ED_OTBANK=" & Val(rsTA("OT_BANK")) & " WHERE ED_EMPNBR=" & rsTA("OT_EMPNBR")
            Else
                SQLQ = "UPDATE HREMP SET ED_OTBANK=" & Val(rsTA("OT_BANK")) - Val(rsTA("OT_BANKT")) & " WHERE ED_EMPNBR=" & rsTA("OT_EMPNBR")
            End If
            gdbAdoIhr001.Execute SQLQ
        End If
        rsTA.MoveNext
    Loop
End If
rsTA.Close


Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodPercent = 100
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""
Exit Sub

ReCalcOvt_Err:
glbFrmCaption$ = "Overtime Bank Recalculation"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ReCalcOvt", "", "Overtime Bank Recalculation")
If gintRollBack% = False Then
    Resume Next
End If

End Sub

Public Function Overtime_Multiplier(xUnion, xCategory, xStatus, xLoc, xRegion, xAdminBy, xSECTION)
Dim rsOvtMst As New ADODB.Recordset
Dim SQLQ As String

    SQLQ = "SELECT OM_MULTIPLIER FROM HR_OVERTIME_MASTER WHERE 1 = 1 "
    If Len(xUnion) > 0 Then
        SQLQ = SQLQ & " AND OM_ORG = '" & xUnion & "'"
    End If
    If Len(xStatus) > 0 Then
        SQLQ = SQLQ & " AND OM_EMP = '" & xStatus & "'"
    End If
    If Len(xCategory) > 0 Then
        SQLQ = SQLQ & " AND OM_PT = '" & xCategory & "'"
    End If
    
    'Ticket #15753
    If Len(xLoc) > 0 Then
        SQLQ = SQLQ & " AND OM_LOC = '" & xLoc & "'"
    End If
    If Len(xRegion) > 0 Then
        SQLQ = SQLQ & " AND OM_REGION = '" & xRegion & "'"
    End If
    If Len(xAdminBy) > 0 Then
        SQLQ = SQLQ & " AND OM_ADMINBY = '" & xAdminBy & "'"
    End If
    If Len(xSECTION) > 0 Then
        SQLQ = SQLQ & " AND OM_SECTION = '" & xSECTION & "'"
    End If
    
    rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsOvtMst.EOF Then
        Overtime_Multiplier = rsOvtMst("OM_MULTIPLIER")
    Else
        Overtime_Multiplier = 0
    End If
    
    rsOvtMst.Close
End Function

Public Function Get_OvertimeBank(xEmpNbr, xFromDate, xToDate)
Dim rsATT As New ADODB.Recordset
Dim rsOvtBank As New ADODB.Recordset
Dim SQLQ As String

    If xFromDate <> "" Then
        SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE WHERE "
        SQLQ = SQLQ & " AD_EMPNBR = " & xEmpNbr & " AND AD_REASON LIKE 'OT%'"
        'SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(Format("1/1/" & Year(Now()), "mm/dd/yyyy")) & " AND AD_DOA <= " & Date_SQL(Format("12/31/" & Year(Now()), "mm/dd/yyyy")) & ""
        SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(xFromDate) & " AND AD_DOA <= " & Date_SQL(xToDate) & ""
    Else
        SQLQ = "SELECT OT_EMPNBR, OT_EFDATE, OT_ETDATE FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & xEmpNbr
        rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsOvtBank.EOF Then
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE WHERE "
            SQLQ = SQLQ & " AD_EMPNBR = " & xEmpNbr & " AND AD_REASON LIKE 'OT%'"
            If Not IsNull(rsOvtBank("OT_EFDATE")) And Not IsNull(rsOvtBank("OT_ETDATE")) Then
                SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsOvtBank("OT_EFDATE")) & " AND AD_DOA <= " & Date_SQL(rsOvtBank("OT_ETDATE")) & ""
            End If
        Else
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE WHERE "
            SQLQ = SQLQ & " AD_EMPNBR = " & xEmpNbr & " AND AD_REASON LIKE 'OT%'"
            'Ticket #18668 - not for calendar year but all that is in the HR_ATTENDANCE
            'SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(Format("1/1/" & Year(Now()), "mm/dd/yyyy")) & " AND AD_DOA <= " & Date_SQL(Format("12/31/" & Year(Now()), "mm/dd/yyyy")) & ""
        End If
        rsOvtBank.Close
    End If
    rsATT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsATT.EOF Then
        If IsNull(rsATT("TOT_HRS")) Then
            Get_OvertimeBank = 0
        Else
            Get_OvertimeBank = rsATT("TOT_HRS")
        End If
    Else
        Get_OvertimeBank = 0
    End If
    rsATT.Close
    
End Function

Public Function Get_OvertimeTaken(xEmpNbr, xFromDate, xToDate)
Dim rsATT As New ADODB.Recordset
Dim rsOvtBank As New ADODB.Recordset
Dim SQLQ As String

    If xFromDate <> "" Then
        SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE WHERE "
        SQLQ = SQLQ & " AD_EMPNBR = " & xEmpNbr & " AND AD_REASON LIKE 'CT%'"
        'SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(Format("1/1/" & Year(Now()), "mm/dd/yyyy")) & " AND AD_DOA <= " & Date_SQL(Format("12/31/" & Year(Now()), "mm/dd/yyyy")) & ""
        SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(xFromDate) & " AND AD_DOA <= " & Date_SQL(xToDate) & ""
    Else
        SQLQ = "SELECT OT_EMPNBR, OT_EFDATE, OT_ETDATE FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & xEmpNbr
        rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsOvtBank.EOF Then
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE WHERE "
            SQLQ = SQLQ & " AD_EMPNBR = " & xEmpNbr & " AND AD_REASON LIKE 'CT%'"
            If Not IsNull(rsOvtBank("OT_EFDATE")) And Not IsNull(rsOvtBank("OT_ETDATE")) Then
                SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsOvtBank("OT_EFDATE")) & " AND AD_DOA <= " & Date_SQL(rsOvtBank("OT_ETDATE")) & ""
            End If
        Else
            SQLQ = "SELECT SUM(AD_HRS) AS TOT_HRS FROM HR_ATTENDANCE WHERE "
            SQLQ = SQLQ & " AD_EMPNBR = " & xEmpNbr & " AND AD_REASON LIKE 'CT%'"
            'Ticket #18668 - not for calendar year but all that is in the HR_ATTENDANCE
            'SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(Format("1/1/" & Year(Now()), "mm/dd/yyyy")) & " AND AD_DOA <= " & Date_SQL(Format("12/31/" & Year(Now()), "mm/dd/yyyy")) & ""
        End If
        rsOvtBank.Close
    End If
    rsATT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not rsATT.EOF Then
        If IsNull(rsATT("TOT_HRS")) Then
            Get_OvertimeTaken = 0
        Else
            Get_OvertimeTaken = rsATT("TOT_HRS")
        End If
    Else
        Get_OvertimeTaken = 0
    End If
    rsATT.Close
    
End Function

Public Function Get_VacationTaken(xEmpNbr, xFromDate, xToDate)
    Dim rzAttend As New ADODB.Recordset
    Dim SQLQ As String
    
    Get_VacationTaken = 0
    
    If xEmpNbr <> "" Then
        SQLQ = "SELECT Sum(AD_HRS) AS SumHRS FROM HR_ATTENDANCE "
        SQLQ = SQLQ & " WHERE LEFT(AD_REASON,3)='VAC' "
        If xFromDate <> "" Then
            SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(xFromDate)
        End If
        If xToDate <> "" Then
            SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(xToDate)
        End If
        SQLQ = SQLQ & " AND AD_EMPNBR = " & xEmpNbr
        rzAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset
        If Not rzAttend.EOF Then
            Get_VacationTaken = rzAttend("SUMHRS")
        End If
        rzAttend.Close
        Set rzAttend = Nothing
    End If
End Function

Public Sub ReCalcUSB(xStr, xFLAG) 'xEmpNo)
Dim rsUSB As New ADODB.Recordset
Dim rsUSBT As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ, xTaken, xRepaid, xOutStand, xHrsDiff
Dim xUnion, xFDate, xTDate, xUnionList

If xFLAG = "EMP" Then
    'Get Union Code from Attendance table
    SQLQ = "SELECT AD_ORG AS ORGCOED FROM HR_ATTENDANCE WHERE " & xStr & " "
    SQLQ = SQLQ & "AND AD_REASON = 'USB' "
    SQLQ = SQLQ & "GROUP BY AD_ORG"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xUnionList = ","
    Screen.MousePointer = HOURGLASS
    Do While Not rsEmp.EOF
        xUnion = ""
        If Not IsNull(rsEmp("ORGCOED")) Then
            xUnion = rsEmp("ORGCOED")
            If InStr(1, xUnionList, xUnion) > 0 Then
                GoTo Next_Line
            End If
            xUnionList = xUnionList & xUnion & ","
        End If

        If Len(xUnion) = 0 Then
            GoTo Next_Line
        End If
        
        SQLQ = "SELECT * FROM WHSCC_USB WHERE WU_ORG = '" & xUnion & "' "
        rsUSB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsUSB.EOF
            xFDate = rsUSB("WU_EFDATE")
            xTDate = rsUSB("WU_ETDATE")
            SQLQ = "SELECT SUM(AD_HRS) AS USBTAKEN FROM HR_ATTENDANCE "
            SQLQ = SQLQ & "WHERE AD_ORG = '" & xUnion & "' "
            SQLQ = SQLQ & "AND AD_REASON = 'USB' "
            SQLQ = SQLQ & "AND AD_DOA >= ('" & Format(xFDate, "mmm dd,yyyy") & "') "
            SQLQ = SQLQ & "AND AD_DOA <= ('" & Format(xTDate, "mmm dd,yyyy") & "') "
            SQLQ = SQLQ & "GROUP BY AD_ORG "
            rsUSBT.Open SQLQ, gdbAdoIhr001, adOpenStatic
            xTaken = 0
            If Not rsUSBT.EOF Then
                xTaken = rsUSBT("USBTAKEN")
            End If
            rsUSBT.Close
            rsUSB("WU_USBT") = xTaken
            rsUSB.Update
            rsUSB.MoveNext
        Loop
        rsUSB.Close
Next_Line:
        rsEmp.MoveNext
    Loop

    rsEmp.Close
    Screen.MousePointer = DEFAULT
End If

If xFLAG = "ORG" Then
        Screen.MousePointer = HOURGLASS
        SQLQ = "SELECT * FROM WHSCC_USB " 'WHERE WU_ORG = '" & xUnion & "' "
        rsUSB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsUSB.EOF
            xUnion = rsUSB("WU_ORG")
            xFDate = rsUSB("WU_EFDATE")
            xTDate = rsUSB("WU_ETDATE")
            SQLQ = "SELECT SUM(AD_HRS) AS USBTAKEN FROM HR_ATTENDANCE "
            SQLQ = SQLQ & "WHERE AD_ORG = '" & xUnion & "' "
            SQLQ = SQLQ & "AND AD_REASON = 'USB' "
            SQLQ = SQLQ & "AND AD_DOA >= ('" & Format(xFDate, "mmm dd,yyyy") & "') "
            SQLQ = SQLQ & "AND AD_DOA <= ('" & Format(xTDate, "mmm dd,yyyy") & "') "
            SQLQ = SQLQ & "GROUP BY AD_ORG "
            rsUSBT.Open SQLQ, gdbAdoIhr001, adOpenStatic
            xTaken = 0
            If Not rsUSBT.EOF Then
                xTaken = rsUSBT("USBTAKEN")
            End If
            rsUSBT.Close
            rsUSB("WU_USBT") = xTaken
            rsUSB.Update
            rsUSB.MoveNext
        Loop
        rsUSB.Close
        Screen.MousePointer = DEFAULT
End If
End Sub

Public Sub ReCalcASL(xEmpNo, xStr)
Dim rsASL As New ADODB.Recordset
Dim rsENT As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ, xTaken, xRepaid, xOutStand, xTotOuts, xID


    If Len(xStr) = 0 Then
        SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES FROM HREMP "
        SQLQ = SQLQ & "WHERE ED_EMPNBR = " & xEmpNo
    Else
        SQLQ = "SELECT ED_EMPNBR,ED_EFDATES,ED_ETDATES " 'FROM HREMP "
        SQLQ = SQLQ & "FROM WHSCC_ASL INNER JOIN HREMP ON WHSCC_ASL.AS_EMPNBR = HREMP.ED_EMPNBR "
        SQLQ = SQLQ & xStr
    End If

        rsENT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsENT.EOF
            xEmpNo = rsENT("ED_EMPNBR")
            If IsDate(rsENT("ED_EFDATES")) And IsDate(rsENT("ED_ETDATES")) Then

                    SQLQ = "SELECT * FROM WHSCC_ASL WHERE AS_EMPNBR = " & xEmpNo & " "
                    'Don't check Date Range for ASL T#3304
                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                    SQLQ = SQLQ & "ORDER BY AS_EMPNBR,AS_DOA "
                    rsASL.CursorLocation = adUseClient
                    rsASL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

                    xTaken = 0: xRepaid = 0: xOutStand = 0: xTotOuts = 0

                    Do While Not rsASL.EOF
                        If IsNull(rsASL("AS_HRSTAK")) Then
                            xTaken = 0
                        Else
                            xTaken = rsASL("AS_HRSTAK")
                        End If
                        If IsNull(rsASL("AS_HRSREP")) Then
                            xRepaid = 0
                        Else
                            xRepaid = rsASL("AS_HRSREP")
                        End If
                        xTotOuts = xTaken - xRepaid
                        xOutStand = xOutStand + xTotOuts
                        xID = rsASL("AS_ATT_ID")
                        SQLQ = "UPDATE WHSCC_ASL SET AS_HRSOS = " & xOutStand & " "
                        SQLQ = SQLQ & "WHERE AS_ATT_ID = " & xID & " "
                        gdbAdoIhr001.Execute SQLQ
                        rsASL.MoveNext
                    Loop
                    rsASL.Close

                    SQLQ = "UPDATE WHSCC_ASL SET AS_HRAOS = " & xOutStand & " "
                    SQLQ = SQLQ & "WHERE AS_EMPNBR = " & xEmpNo & " "
                    'SQLQ = SQLQ & "AND AS_EFDATES = ('" & Format(rsENT("ED_EFDATES"), "mmm dd,yyyy") & "') "
                    'SQLQ = SQLQ & "AND AS_ETDATES = ('" & Format(rsENT("ED_ETDATES"), "mmm dd,yyyy") & "') "
                    gdbAdoIhr001.Execute SQLQ

            End If

            rsENT.MoveNext
        Loop
        rsENT.Close
        Call Pause(0.5)

End Sub

Function ExistTable(db As ADODB.Connection, TableName As String)
Dim rsTT As New ADODB.Recordset
On Error GoTo Error_E
rsTT.Open TableName, db, adOpenForwardOnly, adLockReadOnly, adCmdTableDirect
ExistTable = True

Exit Function

Error_E:
    ExistTable = False
End Function

Sub updateBenefit(xEmpNbr, NewBGroup, TermOrActive, BUpdSource As BenefitUpdateSource, Optional BCodeCover)
Dim rsBGMST As New ADODB.Recordset
Dim rsBN As New ADODB.Recordset
Dim SQLQ As String
Dim xACT
Dim xCode As String, xCover As String
Dim BCSQLQ
Dim xProcessDate
Dim xIfDatChg As Boolean 'Ticket #23729 Franks 05/10/2013
Dim xPer
Dim xDateAge65
Dim xBenType As String

If glbWFC Then 'Ticket #23247 Franks 04/25/2013
    If IsWFCUSBenEmp(xEmpNbr) Then
        Call WFC_UptUSBenByEmp(xEmpNbr, CVDate(Date), 0, "Y", "Y")
        Exit Sub
    End If
End If

'static recordsets changed to optimistic locking
'Bryan 19/Sep/05 Ticket #9327
xProcessDate = getProcessDate(xEmpNbr)
BCSQLQ = ""

If Not IsMissing(BCodeCover) Then
    xCode = Left(BCodeCover, InStr(BCodeCover, "_") - 1)
    xCover = Mid(BCodeCover, InStr(BCodeCover, "_") + 1)
    BCSQLQ = " AND BM_BCODE='" & xCode & "'"
    If Len(xCover) = 0 Then
        BCSQLQ = BCSQLQ & " AND (BM_COVER='" & xCover & "' OR BM_COVER IS NULL) "
    Else
        BCSQLQ = BCSQLQ & " AND BM_COVER='" & xCover & "'"
    End If
Else
    'This is for mass benefit group change. To delete the records that has a group name but not below the current group
    If TermOrActive = "A" Then
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM HRBENFT "
        SQLQ = SQLQ & " LEFT JOIN HR_BENEFITS_GROUP "
        SQLQ = SQLQ & " ON HRBENFT.BF_BCODE=HR_BENEFITS_GROUP.BM_BCODE "
        SQLQ = SQLQ & " AND HRBENFT.BF_COVER=HR_BENEFITS_GROUP.BM_COVER "
        SQLQ = SQLQ & " WHERE BF_GROUP IS NOT  NULL AND BF_EMPNBR IS NULL"
        rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockPessimistic
    Else
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM Term_HRBENFT "
        SQLQ = SQLQ & " LEFT JOIN HR_BENEFITS_GROUP "
        If Not glbSQL And Not glbOracle Then SQLQ = SQLQ & " IN '" & glbIHRDB & "'"
        SQLQ = SQLQ & " ON Term_HRBENFT.BF_BCODE=HR_BENEFITS_GROUP.BM_BCODE "
        SQLQ = SQLQ & " AND Term_HRBENFT.BF_COVER=HR_BENEFITS_GROUP.BM_COVER "
        SQLQ = SQLQ & " WHERE BF_GROUP IS NOT  NULL AND BF_EMPNBR IS NULL "
        
        rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockPessimistic
    End If
    Do Until rsBN.EOF
        rsBN("BF_LUSER") = "999999998"
        rsBN.Update
        If TermOrActive = "A" Then
            Call AUDITBENF("D")
        End If
        rsBN.MoveNext
    Loop
    If TermOrActive = "A" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute "DELETE FROM HRBENFT WHERE BF_LUSER='999999998'"
        gdbAdoIhr001.CommitTrans
    Else
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute "DELETE FROM Term_HRBENFT WHERE BF_LUSER='999999998'"
        gdbAdoIhr001X.CommitTrans
    End If
    rsBN.Close
    Set rsBN = Nothing
End If

SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & NewBGroup & "'"
SQLQ = SQLQ & BCSQLQ
SQLQ = SQLQ & " ORDER BY BM_BCODE "
rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsBGMST.EOF Then
    'This is for Deleting a record from Group Master.
    'It can also be called from employee benefit and the benefit do not exist in the group master
    
    If TermOrActive = "A" Then
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND  BF_GROUP='" & NewBGroup & "'"
        SQLQ = SQLQ & Replace(BCSQLQ, "BM_", "BF_")
        rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockPessimistic
    Else
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM Term_HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND  BF_GROUP='" & NewBGroup & "'"
        SQLQ = SQLQ & Replace(BCSQLQ, "BM_", "BF_")
        rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockPessimistic
    End If
    Do Until rsBN.EOF
        ''Ticket #22243 Franks 07/05/2012 - Jerry doesn't want to detele it, just skip it
        ''rsBN("BF_LUSER") = "999999998"
        ''rsBN.Update
        ''If TermOrActive = "A" Then
        ''    Call AUDITBENF("D")
        ''End If
        ''rsBN.Delete
        
        'Commenting now because of above - Release 8.1 - For Benefit Email Notification
        'glbBenDeleted = "True"
        
        rsBN.MoveNext
    Loop
    rsBN.Close
    Set rsBN = Nothing

Else
    Do While Not rsBGMST.EOF
        xCode = rsBGMST("BM_BCODE")
    
        If BUpdSource = GroupMasterAdd Then
            If TermOrActive = "A" Then
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
                SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
                rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
            Else
                SQLQ = "SELECT * FROM Term_HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
                SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
                rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
            End If
            If Not rsBN.EOF Then
                If IsNull(rsBGMST("BM_COVER")) Or Len(rsBGMST("BM_COVER")) = 0 Then
                   If Not (IsNull(rsBN("BF_COVER")) Or Len(rsBN("BF_COVER")) = 0) Then
                        GoTo NotAddNewRecord
                   End If
                Else
                   If rsBN("BF_COVER") <> rsBGMST("BM_COVER") Then GoTo NotAddNewRecord
                End If
            End If
            rsBN.Close
            Set rsBN = Nothing
        End If
        If TermOrActive = "A" Then
            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
            SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
            If IsNull(rsBGMST("BM_COVER")) Or Len(rsBGMST("BM_COVER")) = 0 Then
                SQLQ = SQLQ & " AND (BF_COVER IS NULL OR BF_COVER='')"
            Else
                SQLQ = SQLQ & " AND BF_COVER='" & rsBGMST("BM_COVER") & "'"
            End If
            rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        Else
            SQLQ = "SELECT * FROM Term_HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
            SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
            If IsNull(rsBGMST("BM_COVER")) Or Len(rsBGMST("BM_COVER")) = 0 Then
                SQLQ = SQLQ & " AND (BF_COVER IS NULL OR BF_COVER='')"
            Else
                SQLQ = SQLQ & " AND BF_COVER='" & rsBGMST("BM_COVER") & "'"
            End If
            rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
        End If
        
        If rsBN.EOF Then
            If BUpdSource = GroupMasterEdit Or BUpdSource = GroupMasterRecal Then
                GoTo NotAddNewRecord
            Else
                rsBN.AddNew
                xACT = "A"
                
                'Release 8.1 - For Benefit Email Notification
                glbBenAdded = "True"
            End If
        Else
            xACT = "M"
            If glbWFC And xCode = "HCSAA" Then 'Ticket #24866 Franks 01/07/2013
                GoTo NotAddNewRecord 'Change RECALCULATE function to exclude HCSAA benefit code.
            End If
            
            'Release 8.1 - For Benefit Email Notification
            glbBenChanged = "True"
        End If
        
        rsBN("BF_COMPNO") = "001"
        rsBN("BF_EMPNBR") = xEmpNbr
        If glbLambton And glbVadim Then
            rsBN("BF_PAYROLL_ID") = Get_Payroll_ID_For_Benefit(xEmpNbr)
        Else
            rsBN("BF_PAYROLL_ID") = GetEmpData(xEmpNbr, "ED_PAYROLL_ID")
        End If
        rsBN("BF_GROUP") = NewBGroup
        rsBN("BF_BCODE") = rsBGMST("BM_BCODE")
        
        'Jerry asked to add this as when Waiting Period is changed it was not changing
        'Employee's Benefit waiting period
        If xACT = "M" Then
            If IsDate(rsBGMST("BM_EDATE")) Then
                rsBN("BF_EDATE") = rsBGMST("BM_EDATE")
            Else
                If IsNumeric(rsBGMST("BM_WAITPERIOD")) Then
                    'Ticket #25152: Macaulay Child Development Centre - PEN Benefit only
                    If glbCompSerial = "S/N - 2420W" And rsBGMST("BM_BCODE") = "PEN" Then
                        rsBN("BF_EDATE") = CountEDate(xEmpNbr, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), , , rsBGMST("BM_BCODE"))
                    Else
                        rsBN("BF_EDATE") = CountEDate(xEmpNbr, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"))
                    End If
                End If
            End If
            rsBN("BF_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
            rsBN("BF_DWM") = rsBGMST("BM_DWM")
        End If
                        
        If xACT = "A" Or glbOttawaCCAC Then       'for ottawa ccac, see ticket #5474
            If IsDate(rsBGMST("BM_EDATE")) Then
                rsBN("BF_EDATE") = rsBGMST("BM_EDATE")
            Else
                'Ticket #25152: Macaulay Child Development Centre - PEN Benefit only
                If glbCompSerial = "S/N - 2420W" And rsBGMST("BM_BCODE") = "PEN" Then
                    rsBN("BF_EDATE") = CountEDate(xEmpNbr, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), , , rsBGMST("BM_BCODE"))
                Else
                    rsBN("BF_EDATE") = CountEDate(xEmpNbr, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"))
                End If
            End If
            rsBN("BF_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
            rsBN("BF_DWM") = rsBGMST("BM_DWM")
        End If
        
        'Release 8.1 - For Email Sending
        glbBenEffDate = rsBN("BF_EDATE")
        
        'Ticket #23729 Franks 05/10/2013 - begin
        xIfDatChg = False
        If xACT = "M" Then
            If Not rsBN("BF_AMT") = rsBGMST("BM_AMT") Then xIfDatChg = True
            If Not rsBN("BF_PPAMT") = rsBGMST("BM_PPAMT") Then xIfDatChg = True
            If Not rsBN("BF_TCOST") = rsBGMST("BM_TCOST") Then xIfDatChg = True
        End If
        'Ticket #23729 Franks 05/10/2013 - end
        rsBN("BF_COVER") = rsBGMST("BM_COVER")
        rsBN("BF_AMT") = rsBGMST("BM_AMT")
        rsBN("BF_PPAMT") = rsBGMST("BM_PPAMT")
        
        'Ticket #25500 - Goodmans - Unit Cost/Rate from Benefits Rate table
        'If glbCompSerial = "S/N - 2290W" And (rsBN("BF_BCODE") = "LIFE" Or rsBN("BF_BCODE") = "SLIFE" Or rsBN("BF_BCODE") = "CLIFE" Or rsBN("BF_BCODE") = "OLIFE") Then
        'If glbCompSerial = "S/N - 2290W" And (rsBN("BF_BCODE") = "SLIFE" Or rsBN("BF_BCODE") = "OLIFE") Then
        'Ticket #27113 - Making option to have different types of Benefit Code setup under Benefit Rates table
        xBenType = ""
        xBenType = Get_BenefitType_BenefitRateTable(rsBN("BF_BCODE"))
        If glbCompSerial = "S/N - 2290W" Then
            'If Left(rsBN("BF_BCODE"), 1) = "S" Then
            '    rsBN("BF_UNITCOST") = Get_BenefitRate(xEmpnbr, rsBN("BF_BCODE"), Spouse)
            'ElseIf Left(rsBN("BF_BCODE"), 1) = "C" Then
            '    rsBN("BF_UNITCOST") = Get_BenefitRate(xEmpnbr, rsBN("BF_BCODE"), Children)
            'ElseIf Left(rsBN("BF_BCODE"), 1) = "O" Then
            '    rsBN("BF_UNITCOST") = Get_BenefitRate(xEmpnbr, rsBN("BF_BCODE"), DependentRelationship.Employee)
            'End If
            If xBenType = "S" Then
                rsBN("BF_UNITCOST") = Get_BenefitRate(xEmpNbr, rsBN("BF_BCODE"), Spouse)
            ElseIf xBenType = "O" Then
                rsBN("BF_UNITCOST") = Get_BenefitRate(xEmpNbr, rsBN("BF_BCODE"), Children)
            ElseIf xBenType = "E" Then
                rsBN("BF_UNITCOST") = Get_BenefitRate(xEmpNbr, rsBN("BF_BCODE"), DependentRelationship.Employee)
            Else
                rsBN("BF_UNITCOST") = rsBGMST("BM_UNITCOST")
            End If
        Else
            rsBN("BF_UNITCOST") = rsBGMST("BM_UNITCOST")
        End If
        
        rsBN("BF_PCE") = rsBGMST("BM_PCE")
        rsBN("BF_PCC") = rsBGMST("BM_PCC")
        If rsBN("BF_SALARYDEPENDANT") = "Y" Then 'Ticket #23729 Franks 05/13/2013
            'do not update these 4 fields
        Else
            rsBN("BF_ECOST") = rsBGMST("BM_ECOST")
            rsBN("BF_CCOST") = rsBGMST("BM_CCOST")
            rsBN("BF_MTHCCOST") = rsBGMST("BM_MTHCCOST")
            rsBN("BF_MTHECOST") = rsBGMST("BM_MTHECOST")
        End If
        rsBN("BF_TCOST") = rsBGMST("BM_TCOST")
        rsBN("BF_MAXDOL") = rsBGMST("BM_MAXDOL")
        rsBN("BF_PREMIUM") = rsBGMST("BM_PREMIUM")
        rsBN("BF_PER") = rsBGMST("BM_PER")
        rsBN("BF_TAXBEN") = rsBGMST("BM_TAXBEN")
        rsBN("BF_SALARYDEPENDANT") = rsBGMST("BM_SALARYDEPENDANT")
        rsBN("BF_MINIMUM") = rsBGMST("BM_MINIMUM")
        rsBN("BF_FACTOR") = rsBGMST("BM_FACTOR")
        rsBN("BF_ROUND") = rsBGMST("BM_ROUND")
        rsBN("BF_MAXIMUM") = rsBGMST("BM_MAXIMUM")
        rsBN("BF_NEXTNEAREST") = rsBGMST("BM_NEXTNEAREST")
        rsBN("BF_TAXAMOUNT") = rsBGMST("BM_TAXAMOUNT")

        rsBN("BF_COMMENTS") = rsBGMST("BM_COMMENTS")
        rsBN("BF_PTAX") = rsBGMST("BM_PTAX")
        rsBN("BF_PERORDOLL") = rsBGMST("BM_PERORDOLL")
        rsBN("BF_POLICY") = rsBGMST("BM_POLICY") 'Ticket #13448 WFC Manulife needs Policy Number
        
        'Ticket #20931 - Rate Level
        rsBN("BF_RATELEVEL") = rsBGMST("BM_RATELEVEL")
        
        'Ticket #23795 - Town of Lasalle - Custom logic to compute Pay Period Amount
        If glbCompSerial = "S/N - 2379W" Then
            'BF_PCC - Company % or BF_PCE - Employee %, and BF_CCOST <> 0 and BF_ECOST <> 0
            If rsBN("BF_PCC") = 1 And rsBN("BF_CCOST") <> 0 Then
                rsBN("BF_PPAMT") = rsBN("BF_CCOST") / 52
                rsBN("BF_PERORDOLL") = "D"
            ElseIf rsBN("BF_PCE") = 1 And rsBN("BF_ECOST") <> 0 Then
                rsBN("BF_PPAMT") = rsBN("BF_ECOST") / 52
                rsBN("BF_PERORDOLL") = "D"
            End If
        End If
        
        If xProcessDate > Date Then
            rsBN("BF_LDATE") = xProcessDate
        Else
            If xACT = "A" Then
                rsBN("BF_LDATE") = rsBN("BF_EDATE")
            Else
                '''The Walter Fedy Partnership - Ticket #15298
                ''If glbCompSerial = "S/N - 2386W" Then
                ''    If CVDate(rsBN("BF_EDATE")) > CVDate(Date) Then
                ''        rsBN("BF_LDATE") = rsBN("BF_EDATE")
                ''    Else
                ''        rsBN("BF_LDATE") = Date
                ''    End If
                ''Else
                ''    rsBN("BF_LDATE") = Date
                ''End If
                'Ticket #28065 Franks 01/29/2016 - SPC got this issue, so make the function for Walter Fedy above for all customer
                If CVDate(rsBN("BF_EDATE")) > CVDate(Date) Then
                    rsBN("BF_LDATE") = rsBN("BF_EDATE")
                Else
                    rsBN("BF_LDATE") = Date
                End If
            End If
        End If
        rsBN("BF_LTIME") = Time$
        
        If TermOrActive = "A" Then
            rsBN("BF_LUSER") = "999999998"
            rsBN.Update
                        
            'Ticket #25500 - Goodmans - LTD Ends Date -> 65th Birthday - 90days -> get the last day of the month
            'If glbCompSerial = "S/N - 2290W" And rsBN("BF_BCODE") = "LTD" Then
            If glbCompSerial = "S/N - 2290W" And rsBN("BF_BCODE") = "LTD" And ((rsBN("BF_GROUP") <> "PARTNERS" And rsBN("BF_GROUP") <> "ART") Or IsNull(rsBN("BF_GROUP"))) Then
                'If IsNumeric(rsBN("BF_WAITPERIOD")) Then
                    'xPER = rsBN("BF_DWM")
                    'If xPER = "W" Then xPER = "ww"
                    
                    'Get the date for Age 65 or 67 based on the Benefit Group
                    If IsNull(rsBN("BF_GROUP")) Or rsBN("BF_GROUP") = "" Then
                        xDateAge65 = DateAdd("yyyy", 67, CVDate(GetEmpData(xEmpNbr, "ED_DOB")))
                    Else
                        xDateAge65 = DateAdd("yyyy", 65, CVDate(GetEmpData(xEmpNbr, "ED_DOB")))
                    End If
                    
                    'Compute LTD End Date based on employee's 65th birthday - 90days and get the last date of month
                    'rsBN("BF_CEASEDATE") = MonthLastDate(DateAdd(xPER, 0 - Val(rsBN("BF_WAITPERIOD")), CVDate(xDateAge65)))
                    'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
                    If IsNull(rsBN("BF_GROUP")) Or rsBN("BF_GROUP") = "" Then
                        'Sept 30th
                        rsBN("BF_CEASEDATE") = CVDate(Format("09/30/" & Year(xDateAge65), "mm/dd/yyyy"))
                    Else
                        rsBN("BF_CEASEDATE") = MonthLastDate(DateAdd("d", 0 - 90, CVDate(xDateAge65)))
                    End If
                    rsBN.Update
                'End If
            End If
                        
            If rsBN("BF_SALARYDEPENDANT") = "Y" Then
                Call updBenefitForSalDEPN(xEmpNbr, True)
            Else
                If xACT = "M" Then 'Ticket #23729 Franks 05/10/2013
                    If xIfDatChg Then
                        Call AUDITBENF(xACT)
                    End If
                Else
                    Call AUDITBENF(xACT)
                End If
            End If
            rsBN.Close
            Set rsBN = Nothing
            rsBN.Open "SELECT BF_LUSER FROM HRBENFT WHERE BF_LUSER='999999998'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        End If
        If Not rsBN.EOF Then
            rsBN("BF_LUSER") = glbUserID
            rsBN.Update
        End If
        If glbMediPay Then 'Ticket #15296
            Call Employee_Benefit_Integration(xEmpNbr)
        End If
NotAddNewRecord:
        rsBN.Close
        Set rsBN = Nothing
        
        rsBGMST.MoveNext
    Loop
    rsBGMST.Close
End If

End Sub

Sub updateBenefit_TERM(xEmpNbr, NewBGroup, TermOrActive, BUpdSource As BenefitUpdateSource, Optional BCodeCover)
Dim rsBGMST As New ADODB.Recordset
Dim rsBN As New ADODB.Recordset
Dim SQLQ As String
Dim xACT
Dim xCode As String, xCover As String
Dim BCSQLQ
Dim xProcessDate
'static recordsets changed to optimistic locking
'Bryan 19/Sep/05 Ticket #9327
xProcessDate = getProcessDate_TERM(xEmpNbr)
BCSQLQ = ""

If Not IsMissing(BCodeCover) Then
    xCode = Left(BCodeCover, InStr(BCodeCover, "_") - 1)
    xCover = Mid(BCodeCover, InStr(BCodeCover, "_") + 1)
    BCSQLQ = " AND BM_BCODE='" & xCode & "'"
    If Len(xCover) = 0 Then
        BCSQLQ = BCSQLQ & " AND (BM_COVER='" & xCover & "' OR BM_COVER IS NULL) "
    Else
        BCSQLQ = BCSQLQ & " AND BM_COVER='" & xCover & "'"
    End If
Else
    'This is for mass benefit group change. To delete the records that has a group name but not below the current group
    If TermOrActive = "A" Then
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM HRBENFT "
        SQLQ = SQLQ & " LEFT JOIN HR_BENEFITS_GROUP "
        SQLQ = SQLQ & " ON HRBENFT.BF_BCODE=HR_BENEFITS_GROUP.BM_BCODE "
        SQLQ = SQLQ & " AND HRBENFT.BF_COVER=HR_BENEFITS_GROUP.BM_COVER "
        SQLQ = SQLQ & " WHERE BF_GROUP IS NOT  NULL AND BF_EMPNBR IS NULL"
        rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockPessimistic
    Else
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM Term_HRBENFT "
        SQLQ = SQLQ & " LEFT JOIN HR_BENEFITS_GROUP "
        If Not glbSQL And Not glbOracle Then SQLQ = SQLQ & " IN '" & glbIHRDB & "'"
        SQLQ = SQLQ & " ON Term_HRBENFT.BF_BCODE=HR_BENEFITS_GROUP.BM_BCODE "
        SQLQ = SQLQ & " AND Term_HRBENFT.BF_COVER=HR_BENEFITS_GROUP.BM_COVER "
        SQLQ = SQLQ & " WHERE BF_GROUP IS NOT  NULL AND BF_EMPNBR IS NULL "
        
        rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockPessimistic
    End If
    Do Until rsBN.EOF
        rsBN("BF_LUSER") = "999999998"
        rsBN.Update
        If TermOrActive = "A" Then
            Call AUDITBENF("D")
        End If
        rsBN.MoveNext
    Loop
    If TermOrActive = "A" Then
        gdbAdoIhr001.BeginTrans
        gdbAdoIhr001.Execute "DELETE FROM HRBENFT WHERE BF_LUSER='999999998'"
        gdbAdoIhr001.CommitTrans
    Else
        gdbAdoIhr001X.BeginTrans
        gdbAdoIhr001X.Execute "DELETE FROM Term_HRBENFT WHERE BF_LUSER='999999998'"
        gdbAdoIhr001X.CommitTrans
    End If
    rsBN.Close
End If

SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & NewBGroup & "'"
SQLQ = SQLQ & BCSQLQ
SQLQ = SQLQ & " ORDER BY BM_BCODE "
rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsBGMST.EOF Then
    'This is for Deleting a record from Group Master.
    'It can also be called from employee benefit and the benefit do not exist in the group master
    
    If TermOrActive = "A" Then
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND  BF_GROUP='" & NewBGroup & "'"
        SQLQ = SQLQ & Replace(BCSQLQ, "BM_", "BF_")
        rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockPessimistic
    Else
        SQLQ = "SELECT BF_BENE_ID,BF_LUSER FROM Term_HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND  BF_GROUP='" & NewBGroup & "'"
        SQLQ = SQLQ & Replace(BCSQLQ, "BM_", "BF_")
        rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockPessimistic
    End If
    Do Until rsBN.EOF
        rsBN("BF_LUSER") = "999999998"
        rsBN.Update
        If TermOrActive = "A" Then
            Call AUDITBENF("D")
        End If
        rsBN.Delete
        rsBN.MoveNext
    Loop
    rsBN.Close
Else
    Do While Not rsBGMST.EOF
        xCode = rsBGMST("BM_BCODE")
    
        If BUpdSource = GroupMasterAdd Then
            If TermOrActive = "A" Then
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
                SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
                rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
            Else
                SQLQ = "SELECT * FROM Term_HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
                SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
                rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
            End If
            If Not rsBN.EOF Then
                If IsNull(rsBGMST("BM_COVER")) Or Len(rsBGMST("BM_COVER")) = 0 Then
                   If Not (IsNull(rsBN("BF_COVER")) Or Len(rsBN("BF_COVER")) = 0) Then
                        GoTo NotAddNewRecord
                   End If
                Else
                   If rsBN("BF_COVER") <> rsBGMST("BM_COVER") Then GoTo NotAddNewRecord
                End If
            End If
            rsBN.Close
        End If
        If TermOrActive = "A" Then
            SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
            SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
            If IsNull(rsBGMST("BM_COVER")) Or Len(rsBGMST("BM_COVER")) = 0 Then
                SQLQ = SQLQ & " AND (BF_COVER IS NULL OR BF_COVER='')"
            Else
                SQLQ = SQLQ & " AND BF_COVER='" & rsBGMST("BM_COVER") & "'"
            End If
            rsBN.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
        Else
            SQLQ = "SELECT * FROM Term_HRBENFT WHERE BF_EMPNBR = " & xEmpNbr
            SQLQ = SQLQ & " AND  BF_BCODE='" & rsBGMST("BM_BCODE") & "'"
            If IsNull(rsBGMST("BM_COVER")) Or Len(rsBGMST("BM_COVER")) = 0 Then
                SQLQ = SQLQ & " AND (BF_COVER IS NULL OR BF_COVER='')"
            Else
                SQLQ = SQLQ & " AND BF_COVER='" & rsBGMST("BM_COVER") & "'"
            End If
            rsBN.Open SQLQ, gdbAdoIhr001X, adOpenDynamic, adLockOptimistic
        End If
        
        If rsBN.EOF Then
            If BUpdSource = GroupMasterEdit Or BUpdSource = GroupMasterRecal Then
                GoTo NotAddNewRecord
            Else
                rsBN.AddNew
                xACT = "A"
            End If
        Else
            xACT = "M"
        End If
        
        rsBN("BF_COMPNO") = "001"
        rsBN("BF_EMPNBR") = xEmpNbr
        If glbLambton And glbVadim Then
            rsBN("BF_PAYROLL_ID") = Get_Payroll_ID_For_Benefit(xEmpNbr)
        Else
            rsBN("BF_PAYROLL_ID") = GetEmpData(xEmpNbr, "ED_PAYROLL_ID")
        End If
        rsBN("BF_GROUP") = NewBGroup
        rsBN("BF_BCODE") = rsBGMST("BM_BCODE")
        
        'Jerry asked to add this as when Waiting Period is changed it was not changing
        'Employee's Benefit waiting period
        If xACT = "M" Then
            If IsDate(rsBGMST("BM_EDATE")) Then
                rsBN("BF_EDATE") = rsBGMST("BM_EDATE")
            Else
                If IsNumeric(rsBGMST("BM_WAITPERIOD")) Then
                    rsBN("BF_EDATE") = CountEDate(xEmpNbr, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), "T")
                End If
            End If
            rsBN("BF_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
            rsBN("BF_DWM") = rsBGMST("BM_DWM")
        End If
        
        If xACT = "A" Or glbOttawaCCAC Then       'for ottawa ccac, see ticket #5474
            If IsDate(rsBGMST("BM_EDATE")) Then
                rsBN("BF_EDATE") = rsBGMST("BM_EDATE")
            Else
                rsBN("BF_EDATE") = CountEDate(xEmpNbr, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"), "T")
            End If
            rsBN("BF_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
            rsBN("BF_DWM") = rsBGMST("BM_DWM")
        End If
        
        rsBN("BF_COVER") = rsBGMST("BM_COVER")
        rsBN("BF_AMT") = rsBGMST("BM_AMT")
        rsBN("BF_PPAMT") = rsBGMST("BM_PPAMT")
        rsBN("BF_UNITCOST") = rsBGMST("BM_UNITCOST")
        rsBN("BF_PCE") = rsBGMST("BM_PCE")
        rsBN("BF_PCC") = rsBGMST("BM_PCC")
        rsBN("BF_ECOST") = rsBGMST("BM_ECOST")
        rsBN("BF_CCOST") = rsBGMST("BM_CCOST")
        rsBN("BF_TCOST") = rsBGMST("BM_TCOST")
        rsBN("BF_MAXDOL") = rsBGMST("BM_MAXDOL")
        rsBN("BF_PREMIUM") = rsBGMST("BM_PREMIUM")
        rsBN("BF_PER") = rsBGMST("BM_PER")
        rsBN("BF_MTHCCOST") = rsBGMST("BM_MTHCCOST")
        rsBN("BF_MTHECOST") = rsBGMST("BM_MTHECOST")
        rsBN("BF_TAXBEN") = rsBGMST("BM_TAXBEN")
        rsBN("BF_SALARYDEPENDANT") = rsBGMST("BM_SALARYDEPENDANT")
        rsBN("BF_MINIMUM") = rsBGMST("BM_MINIMUM")
        rsBN("BF_FACTOR") = rsBGMST("BM_FACTOR")
        rsBN("BF_ROUND") = rsBGMST("BM_ROUND")
        rsBN("BF_MAXIMUM") = rsBGMST("BM_MAXIMUM")
        rsBN("BF_NEXTNEAREST") = rsBGMST("BM_NEXTNEAREST")
        rsBN("BF_TAXAMOUNT") = rsBGMST("BM_TAXAMOUNT")

        rsBN("BF_COMMENTS") = rsBGMST("BM_COMMENTS")
        rsBN("BF_PTAX") = rsBGMST("BM_PTAX")
        rsBN("BF_PERORDOLL") = rsBGMST("BM_PERORDOLL")
        rsBN("BF_POLICY") = rsBGMST("BM_POLICY") 'Ticket #13448 WFC Manulife needs Policy Number
        
        'Ticket #23795 - Town of Lasalle - Custom logic to compute Pay Period Amount
        If glbCompSerial = "S/N - 2379W" Then
            'BF_PCC - Company % or BF_PCE - Employee %, and BF_CCOST <> 0 and BF_ECOST <> 0
            If rsBN("BF_PCC") = 1 And rsBN("BF_CCOST") <> 0 Then
                rsBN("BF_PPAMT") = rsBN("BF_CCOST") / 52
                rsBN("BF_PERORDOLL") = "D"
            ElseIf rsBN("BF_PCE") = 1 And rsBN("BF_ECOST") <> 0 Then
                rsBN("BF_PPAMT") = rsBN("BF_ECOST") / 52
                rsBN("BF_PERORDOLL") = "D"
            End If
        End If
        
        If xProcessDate > Date Then
            rsBN("BF_LDATE") = xProcessDate
        Else
            If xACT = "A" Then
                rsBN("BF_LDATE") = rsBN("BF_EDATE")
            Else
                '''The Walter Fedy Partnership - Ticket #15298
                ''If glbCompSerial = "S/N - 2386W" Then
                ''    If CVDate(rsBN("BF_EDATE")) > CVDate(Date) Then
                ''        rsBN("BF_LDATE") = rsBN("BF_EDATE")
                ''    Else
                ''        rsBN("BF_LDATE") = Date
                ''    End If
                ''Else
                ''    rsBN("BF_LDATE") = Date
                ''End If
                'Ticket #28065 Franks 01/29/2016 - SPC got this issue, so make the function for Walter Fedy above for all customer
                If CVDate(rsBN("BF_EDATE")) > CVDate(Date) Then
                    rsBN("BF_LDATE") = rsBN("BF_EDATE")
                Else
                    rsBN("BF_LDATE") = Date
                End If
            End If
        End If
        
        rsBN("BF_LTIME") = Time$
        
        If TermOrActive = "A" Then
            rsBN("BF_LUSER") = "999999998"
            rsBN.Update
            If rsBN("BF_SALARYDEPENDANT") = "Y" Then
                Call updBenefitForSalDEPN(xEmpNbr, True)
            Else
                Call AUDITBENF(xACT)
            End If
            rsBN.Close
            rsBN.Open "SELECT BF_LUSER FROM HRBENFT WHERE BF_LUSER='999999998'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        End If
        If Not rsBN.EOF Then
            rsBN("BF_LUSER") = glbUserID
            rsBN.Update
        End If
        'If glbMediPay Then 'Ticket #15296
        '    Call Employee_Benefit_Integration(xEMPNBR)
        'End If
NotAddNewRecord:
        rsBN.Close
        rsBGMST.MoveNext
    Loop
    rsBGMST.Close
End If

End Sub

Private Function AUDITBENF(ACTX)
Dim TA As New ADODB.Recordset
Dim TB As New ADODB.Recordset
Dim xPT, xDiv, xADD
Dim TC As New ADODB.Recordset
Dim SQLQ
On Error GoTo AUDIT_ERR
AUDITBENF = False
If glbSQL Or glbOracle Then
    SQLQ = "INSERT INTO HRAUDIT ("
    SQLQ = SQLQ & " AU_COMPNO"
    SQLQ = SQLQ & ",AU_EMPNBR"
    SQLQ = SQLQ & ",AU_PTUPL"
    SQLQ = SQLQ & ",AU_DIVUPL"
    SQLQ = SQLQ & ",AU_NEWEMP"
    SQLQ = SQLQ & ",AU_BCODE"
    SQLQ = SQLQ & ",AU_COVER"
    SQLQ = SQLQ & ",AU_MAXDOL"
    SQLQ = SQLQ & ",AU_EDATE"
    SQLQ = SQLQ & ",AU_LDATE"
    SQLQ = SQLQ & ",AU_LUSER"
    SQLQ = SQLQ & ",AU_LTIME"
    SQLQ = SQLQ & ",AU_UPLOAD"
    SQLQ = SQLQ & ",AU_TYPE"
    SQLQ = SQLQ & ",AU_TCOST"
    SQLQ = SQLQ & ",AU_PREMIUM"
    SQLQ = SQLQ & ",AU_PCE"
    SQLQ = SQLQ & ",AU_PCC"
    SQLQ = SQLQ & ",AU_PPAMT"
    SQLQ = SQLQ & ",AU_PER"
    SQLQ = SQLQ & ",AU_BAMT"
    SQLQ = SQLQ & ",AU_UNITCOST"
    SQLQ = SQLQ & ",AU_MTHECOST"
    SQLQ = SQLQ & ",AU_MTHCCOST"
    SQLQ = SQLQ & ",AU_PAYROLL_ID"
    SQLQ = SQLQ & " )"
    SQLQ = SQLQ & " SELECT"
    SQLQ = SQLQ & " '001'"
    SQLQ = SQLQ & ",BF_EMPNBR"
    SQLQ = SQLQ & ",ED_PT"
    SQLQ = SQLQ & ",ED_DIV"
    SQLQ = SQLQ & ",'N'"
    SQLQ = SQLQ & ",BF_BCODE"
    SQLQ = SQLQ & ",BF_COVER"
    SQLQ = SQLQ & ",BF_MAXDOL"
    SQLQ = SQLQ & ",BF_EDATE"
    ''Ticket #20843 Franks 08/23/2011 - use TODAY as AU_LDATE
    ''SQLQ = SQLQ & ",BF_LDATE"
    'Ticket #21960 Franks 04/26/2012
    'SQLQ = SQLQ & "," & Date_SQL(Date) & ""
    SQLQ = SQLQ & ",(SELECT (CASE WHEN (BF_LDATE > GETDATE()) THEN BF_LDATE ELSE DATEDIFF(DD,0,GETDATE()) END))"
    SQLQ = SQLQ & ",'" & glbUserID & "'"
    SQLQ = SQLQ & ",BF_LTIME"
    SQLQ = SQLQ & ",'N'"
    SQLQ = SQLQ & ",'" & ACTX & "'"
    SQLQ = SQLQ & ",BF_TCOST"
    SQLQ = SQLQ & ",BF_PREMIUM"
    SQLQ = SQLQ & ",BF_PCE"
    SQLQ = SQLQ & ",BF_PCC"
    SQLQ = SQLQ & ",BF_PPAMT"
    SQLQ = SQLQ & ",BF_PER"
    SQLQ = SQLQ & ",BF_AMT"
    SQLQ = SQLQ & ",BF_UNITCOST"
    SQLQ = SQLQ & ",BF_MTHECOST"
    SQLQ = SQLQ & ",BF_MTHCCOST"
    SQLQ = SQLQ & ",ED_PAYROLL_ID"
    
    If glbOracle Then
        SQLQ = SQLQ & " FROM HRBENFT, HREMP WHERE HRBENFT.BF_EMPNBR=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " AND BF_LUSER='999999998' "
    Else
        SQLQ = SQLQ & " FROM HRBENFT INNER JOIN HREMP ON HRBENFT.BF_EMPNBR=HREMP.ED_EMPNBR "
        SQLQ = SQLQ & " WHERE BF_LUSER='999999998' "
    End If
    
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute SQLQ
    gdbAdoIhr001.CommitTrans
Else
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_LUSER='999999998'"
    TC.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    SQLQ = "SELECT "
    SQLQ = SQLQ & " AU_COMPNO"
    SQLQ = SQLQ & ",AU_EMPNBR"
    SQLQ = SQLQ & ",AU_PTUPL"
    SQLQ = SQLQ & ",AU_DIVUPL"
    SQLQ = SQLQ & ",AU_NEWEMP"
    SQLQ = SQLQ & ",AU_BCODE"
    SQLQ = SQLQ & ",AU_COVER"
    SQLQ = SQLQ & ",AU_MAXDOL"
    SQLQ = SQLQ & ",AU_EDATE"
    SQLQ = SQLQ & ",AU_LDATE"
    SQLQ = SQLQ & ",AU_LUSER"
    SQLQ = SQLQ & ",AU_LTIME"
    SQLQ = SQLQ & ",AU_UPLOAD"
    SQLQ = SQLQ & ",AU_TYPE"
    SQLQ = SQLQ & ",AU_TCOST"
    SQLQ = SQLQ & ",AU_PREMIUM"
    SQLQ = SQLQ & ",AU_PCE"
    SQLQ = SQLQ & ",AU_PCC"
    SQLQ = SQLQ & ",AU_PPAMT"
    SQLQ = SQLQ & ",AU_PER"
    SQLQ = SQLQ & ",AU_BAMT"
    SQLQ = SQLQ & ",AU_UNITCOST"
    SQLQ = SQLQ & ",AU_MTHECOST"
    SQLQ = SQLQ & ",AU_MTHCCOST"
    SQLQ = SQLQ & ",AU_PAYROLL_ID"
    SQLQ = SQLQ & " FROM HRAUDIT WHERE AU_EMPNBR=0"
    TA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    xADD = False
    Do Until TC.EOF
        TA.AddNew
        TB.Open "SELECT ED_EMPNBR,ED_PT,ED_DIV,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR=" & TC("BF_EMPNBR"), gdbAdoIhr001, adOpenKeyset
        If Not TB.EOF Then
            TA("AU_PTUPL") = TB("ED_PT")
            TA("AU_DIVUPL") = TB("ED_DIV")
            If Not IsNull(TB("ED_PAYROLL_ID")) Then
                TA("AU_PAYROLL_ID") = TB("ED_PAYROLL_ID")
            End If
        End If
        TB.Close
        'TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR": TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL": TA("AU_EARN_TABL") = "EARN"
        TA("AU_NEWEMP") = "N"
        TA("AU_BCODE") = TC("BF_BCODE")
        TA("AU_COVER") = TC("BF_COVER")
        TA("AU_MAXDOL") = TC("BF_MAXDOL")
        TA("AU_EDATE") = TC("BF_EDATE")

        TA("AU_TCOST") = TC("BF_TCOST")
        TA("AU_PREMIUM") = TC("BF_PREMIUM")
        TA("AU_PCE") = TC("BF_PCE")
        TA("AU_PCC") = TC("BF_PCC")
        TA("AU_PPAMT") = TC("BF_PPAMT")
        TA("AU_PER") = TC("BF_PER")
        TA("AU_BAMT") = TC("BF_AMT")
        TA("AU_UNITCOST") = TC("BF_UNITCOST")
        TA("AU_MTHECOST") = TC("BF_MTHECOST")
        TA("AU_MTHCCOST") = TC("BF_MTHCCOST")
        
        TA("AU_COMPNO") = "001"
        TA("AU_EMPNBR") = TC("BF_EMPNBR")
        TA("AU_LDATE") = Date
        TA("AU_LUSER") = glbUserID
        TA("AU_LTIME") = Time$
        TA("AU_UPLOAD") = "N"
        TA("AU_TYPE") = ACTX
        TA.Update
        TC.MoveNext
    Loop
    TC.Close
End If
AUDITBENF = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = "UpdData"
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")

End Function

Function CountEDate(xEmpNbr, WaitPeriod, DWM, Optional xDOH, Optional xType, Optional xBenCode)
    Dim DWM_Format
    Dim rsEmp As New ADODB.Recordset
    Dim xDate, xYY, xMM, xDD
    CountEDate = Date
    If IsNumeric(WaitPeriod) Then
        If Not IsMissing(xType) Then
            If xType = "T" Then
                rsEmp.Open "SELECT ED_DOH,ED_USRDAT1 FROM TERM_HREMP WHERE ED_EMPNBR=" & xEmpNbr, gdbAdoIhr001X, adOpenStatic
            Else
                rsEmp.Open "SELECT ED_DOH,ED_USRDAT1 FROM HREMP WHERE ED_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenStatic
            End If
        Else
            rsEmp.Open "SELECT ED_DOH,ED_USRDAT1 FROM HREMP WHERE ED_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenStatic
        End If
        If Not rsEmp.EOF Then
            'Ticket #24203 - Family Day Care Services
            'Ticket #21504 - Kerry's Place
            If (glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2436W") And IsDate(rsEmp("ED_USRDAT1")) Then
                If IsNull(DWM) Or Len(DWM) = 0 Then
                    DWM_Format = "m"
                ElseIf DWM = "W" Then
                    DWM_Format = "ww"
                Else
                    DWM_Format = DWM
                End If
                CountEDate = DateAdd(DWM_Format, WaitPeriod, rsEmp("ED_USRDAT1"))
            ElseIf IsDate(rsEmp("ED_DOH")) Then
                If Not glbCElgin Then
                    If IsNull(DWM) Or Len(DWM) = 0 Then
                        DWM_Format = "m"
                    ElseIf DWM = "W" Then
                        DWM_Format = "ww"
                    Else
                        DWM_Format = DWM
                    End If
                    'Ticket #25152: Macaulay Child Development Centre - PEN Benefit only
                    If glbCompSerial = "S/N - 2420W" And Not IsMissing(xBenCode) Then
                        If xBenCode = "PEN" Then
                            If Day(rsEmp("ED_DOH")) = 1 Then
                                CountEDate = DateAdd(DWM_Format, WaitPeriod, rsEmp("ED_DOH"))
                            Else
                                xDate = MonthLastDate(rsEmp("ED_DOH"))
                                CountEDate = DateAdd("d", 1, CVDate(xDate))
                            End If
                        Else
                            CountEDate = DateAdd(DWM_Format, WaitPeriod, rsEmp("ED_DOH"))
                        End If
                    Else
                        CountEDate = DateAdd(DWM_Format, WaitPeriod, rsEmp("ED_DOH"))
                    End If
                Else
                    xDate = DateAdd("d", Val(WaitPeriod), rsEmp("ED_DOH"))
                    xDD = Day(CVDate(xDate))
                    If xDD > 15 Then
                        xDate = DateAdd("d", -(xDD - 1), CVDate(xDate))
                        CountEDate = DateAdd("m", 1, CVDate(xDate))
                    Else
                        CountEDate = CVDate(xDate)
                    End If
                End If
            ElseIf Not IsMissing(xDOH) Then
                If Not glbCElgin Then
                    If IsNull(DWM) Or Len(DWM) = 0 Then
                        DWM_Format = "m"
                    ElseIf DWM = "W" Then
                        DWM_Format = "ww"
                    Else
                        DWM_Format = DWM
                    End If
                    CountEDate = DateAdd(DWM_Format, WaitPeriod, xDOH)
                Else
                    xDate = DateAdd("d", Val(WaitPeriod), xDOH)
                    xDD = Day(CVDate(xDate))
                    If xDD > 15 Then
                        xDate = DateAdd("d", -(xDD - 1), CVDate(xDate))
                        CountEDate = DateAdd("m", 1, CVDate(xDate))
                    Else
                        CountEDate = CVDate(xDate)
                    End If
                End If
            End If
        End If
        rsEmp.Close
    Else
        If glbCompSerial = "S/N - 2436W" Then 'Family Day Ticket #26114 Franks 11/14/2014
            If IsDate(xDOH) Then
                CountEDate = CVDate(xDOH)
            End If
        End If
    End If
End Function

Public Function GetBGData(BGroup, BCode, Field As String, DEFAULT)
    Dim rsBGR As New ADODB.Recordset
    rsBGR.Open "SELECT " & Field & " FROM HR_BENEFITS_GROUP WHERE BM_BCODE='" & BCode & "' AND BM_BENEFIT_GROUP='" & BGroup & "'", gdbAdoIhr001, adOpenForwardOnly
    GetBGData = DEFAULT
    
    If Not rsBGR.EOF Then
        If Not IsNull(rsBGR(Field)) Then GetBGData = rsBGR(Field)
    End If
End Function

Public Function GetJHData(EmpNbr, Field As String, DEFAULT)
    Dim rsJHTEMP As New ADODB.Recordset
    rsJHTEMP.Open "SELECT " & Field & " FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenForwardOnly
    GetJHData = DEFAULT
    
    If Not rsJHTEMP.EOF Then
        If Not IsNull(rsJHTEMP(Field)) Then GetJHData = rsJHTEMP(Field)
    End If
End Function

Public Function GetTermJHData(EmpNbr, Field As String, DEFAULT, Optional xTermSEQ)
    Dim rsJHTEMP As New ADODB.Recordset
    If Not IsMissing(xTermSEQ) Then
        rsJHTEMP.Open "SELECT " & Field & " FROM TERM_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr & " AND TERM_SEQ = " & xTermSEQ, gdbAdoIhr001, adOpenForwardOnly
    Else
        rsJHTEMP.Open "SELECT TOP 1 " & Field & " FROM TERM_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr & " ORDER BY JH_SDATE DESC", gdbAdoIhr001, adOpenForwardOnly
    End If
    GetTermJHData = DEFAULT
    
    If Not rsJHTEMP.EOF Then
        If Not IsNull(rsJHTEMP(Field)) Then GetTermJHData = rsJHTEMP(Field)
    End If
End Function

Public Function GetSHData(EmpNbr, Field As String, DEFAULT)
    Dim rsSHTEMP As New ADODB.Recordset
    rsSHTEMP.Open "SELECT " & Field & " FROM HR_SALARY_HISTORY WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenForwardOnly
    
    GetSHData = DEFAULT
    If Not rsSHTEMP.EOF Then
        If Not IsNull(rsSHTEMP(Field)) Then GetSHData = rsSHTEMP(Field)
    End If
End Function

Public Function GetTermSHData(EmpNbr, Field As String, DEFAULT, Optional xTermSEQ)
    Dim rsSHTEMP As New ADODB.Recordset
    If Not IsMissing(xTermSEQ) Then
        rsSHTEMP.Open "SELECT " & Field & " FROM TERM_SALARY_HISTORY WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EmpNbr & " AND TERM_SEQ = " & xTermSEQ, gdbAdoIhr001, adOpenForwardOnly
    Else
        rsSHTEMP.Open "SELECT TOP 1 " & Field & " FROM TERM_SALARY_HISTORY WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EmpNbr & " ORDER BY SH_EDATE DESC", gdbAdoIhr001, adOpenForwardOnly
    End If
    
    GetTermSHData = DEFAULT
    If Not rsSHTEMP.EOF Then
        If Not IsNull(rsSHTEMP(Field)) Then GetTermSHData = rsSHTEMP(Field)
    End If
End Function

Public Function GetJobData(JobCode, Field As String, Optional DEFAULT) As String
    Dim rsJOB As New ADODB.Recordset
    rsJOB.Open "SELECT " & Field & " FROM HRJOB WHERE JB_CODE='" & JobCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If Not IsMissing(DEFAULT) Then
        GetJobData = DEFAULT
    Else
        GetJobData = ""
    End If
    If Not rsJOB.EOF Then
        If Not IsNull(rsJOB(Field)) Then GetJobData = rsJOB(Field)
    End If
End Function

Public Function Get_Maximum_Bank_Hours(EmpNbr)
    Dim rsEE As New ADODB.Recordset
    Dim rsOvtMst As New ADODB.Recordset
    Dim xUnion As String
    Dim SQLQ As String
    Dim xStatusDate As Boolean
    Dim xWHERE As String
    Dim flgEmp, flgLoc, flgPT, flgRegion, flgSection, flgAdminBy As Boolean
    
    rsEE.Open "SELECT ED_EMPNBR, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
    If glbtermopen And glbVadim Then
        rsEE.Close
        rsEE.Open "SELECT ED_EMPNBR, ED_ORG, ED_PT, ED_EMP, ED_LOC, ED_REGION, ED_ADMINBY, ED_SECTION FROM Term_HREMP WHERE ED_EMPNBR=" & EmpNbr & " ORDER BY TERM_SEQ DESC", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        Get_Maximum_Bank_Hours = ""
    End If
    
    xStatusDate = False
    If Not rsEE.EOF Then
        If UCase(MDIMain.ActiveForm.name) = "FRMEESTATS" Then
            xUnion = frmEESTATS.clpCode(2).Text
            xStatusDate = True
        Else
            xUnion = IIf(IsNull(rsEE("ED_ORG")), "", rsEE("ED_ORG"))
            xStatusDate = False
        End If
        If Not IsNull(xUnion) And xUnion <> "" Then
            SQLQ = "SELECT * FROM HR_OVERTIME_MASTER WHERE OM_ORG = '" & xUnion & "'"
            rsOvtMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsOvtMst.EOF Then
                rsOvtMst.MoveFirst
                
                Do While Not rsOvtMst.EOF
                    'Building IF condition AND operator
                    xWHERE = ""
                    If Not IsNull(rsOvtMst("OM_EMP")) And rsOvtMst("OM_EMP") <> "" Then
                        If xStatusDate Then
                            xWHERE = "('" & frmEESTATS.clpCode(1).Text & "'='" & rsOvtMst("OM_EMP") & "')"
                            If frmEESTATS.clpCode(1).Text = rsOvtMst("OM_EMP") Then
                                flgEmp = True
                            Else
                                flgEmp = False
                            End If
                        Else
                            xWHERE = "('" & rsEE("ED_EMP") & "'='" & rsOvtMst("OM_EMP") & "')"
                            If rsEE("ED_EMP") = rsOvtMst("OM_EMP") Then
                                flgEmp = True
                            Else
                                flgEmp = False
                            End If
                        End If
                    Else
                        flgEmp = True
                    End If
                    If Not IsNull(rsOvtMst("OM_PT")) And rsOvtMst("OM_PT") <> "" Then
                        If Len(xWHERE) > 0 Then
                            If xStatusDate Then
                                xWHERE = xWHERE & " AND " & "('" & frmEESTATS.clpPT.Text & "'='" & rsOvtMst("OM_PT") & "')"
                                If frmEESTATS.clpPT.Text = rsOvtMst("OM_PT") Then
                                    flgPT = True
                                Else
                                    flgPT = False
                                End If
                            Else
                                xWHERE = xWHERE & " AND " & "('" & rsEE("ED_PT") & "'='" & rsOvtMst("OM_PT") & "')"
                                If rsEE("ED_PT") = rsOvtMst("OM_PT") Then
                                    flgPT = True
                                Else
                                    flgPT = False
                                End If
                            End If
                        Else
                            If xStatusDate Then
                                xWHERE = "('" & frmEESTATS.clpPT.Text & "'='" & rsOvtMst("OM_PT") & "')"
                                If frmEESTATS.clpPT.Text = rsOvtMst("OM_PT") Then
                                    flgPT = True
                                Else
                                    flgPT = False
                                End If
                            Else
                                xWHERE = "('" & rsEE("ED_PT") & "'='" & rsOvtMst("OM_PT") & "')"
                                If rsEE("ED_PT") = rsOvtMst("OM_PT") Then
                                    flgPT = True
                                Else
                                    flgPT = False
                                End If
                            End If
                        End If
                    Else
                        flgPT = True
                    End If
                    If Not IsNull(rsOvtMst("OM_LOC")) And rsOvtMst("OM_LOC") <> "" Then
                        If Len(xWHERE) > 0 Then
                            If xStatusDate Then
                                xWHERE = xWHERE & " AND " & "('" & rsEE("ED_LOC") & "'='" & rsOvtMst("OM_LOC") & "')"
                                If rsEE("ED_LOC") = rsOvtMst("OM_LOC") Then
                                    flgLoc = True
                                Else
                                    flgLoc = False
                                End If
                            Else
                                xWHERE = xWHERE & " AND " & "('" & frmEEBASIC.clpCode(1).Text & "'='" & rsOvtMst("OM_LOC") & "')"
                                If frmEEBASIC.clpCode(1).Text = rsOvtMst("OM_LOC") Then
                                    flgLoc = True
                                Else
                                    flgLoc = False
                                End If
                            End If
                        Else
                            If xStatusDate Then
                                xWHERE = "('" & rsEE("ED_LOC") & "'='" & rsOvtMst("OM_LOC") & "')"
                                If rsEE("ED_LOC") = rsOvtMst("OM_LOC") Then
                                    flgLoc = True
                                Else
                                    flgLoc = False
                                End If
                            Else
                                xWHERE = "('" & frmEEBASIC.clpCode(1).Text & "'='" & rsOvtMst("OM_LOC") & "')"
                                If frmEEBASIC.clpCode(1).Text = rsOvtMst("OM_LOC") Then
                                    flgLoc = True
                                Else
                                    flgLoc = False
                                End If
                            End If
                        End If
                    Else
                        flgLoc = True
                    End If
                    If Not IsNull(rsOvtMst("OM_REGION")) And rsOvtMst("OM_REGION") <> "" Then
                        If Len(xWHERE) > 0 Then
                            If xStatusDate Then
                                xWHERE = xWHERE & " AND " & "('" & rsEE("ED_REGION") & "'='" & rsOvtMst("OM_REGION") & "')"
                                If rsEE("ED_REGION") = rsOvtMst("OM_REGION") Then
                                    flgRegion = True
                                Else
                                    flgRegion = False
                                End If
                            Else
                                xWHERE = xWHERE & " AND " & "('" & frmEEBASIC.clpCode(2).Text & "'='" & rsOvtMst("OM_REGION") & "')"
                                If frmEEBASIC.clpCode(2).Text = rsOvtMst("OM_REGION") Then
                                    flgRegion = True
                                Else
                                    flgRegion = False
                                End If
                            End If
                        Else
                            If xStatusDate Then
                                xWHERE = "('" & rsEE("ED_REGION") & "'='" & rsOvtMst("OM_REGION") & "')"
                                If rsEE("ED_REGION") = rsOvtMst("OM_REGION") Then
                                    flgRegion = True
                                Else
                                    flgRegion = False
                                End If
                            Else
                                xWHERE = "('" & frmEEBASIC.clpCode(2).Text & "'='" & rsOvtMst("OM_REGION") & "')"
                                If frmEEBASIC.clpCode(2).Text = rsOvtMst("OM_REGION") Then
                                    flgRegion = True
                                Else
                                    flgRegion = False
                                End If
                            End If
                        End If
                    Else
                        flgRegion = True
                    End If
                    If Not IsNull(rsOvtMst("OM_ADMINBY")) And rsOvtMst("OM_ADMINBY") <> "" Then
                        If Len(xWHERE) > 0 Then
                            If xStatusDate Then
                                xWHERE = xWHERE & " AND " & "('" & rsEE("ED_ADMINBY") & "'='" & rsOvtMst("OM_ADMINBY") & "')"
                                If rsEE("ED_ADMINBY") = rsOvtMst("OM_ADMINBY") Then
                                    flgAdminBy = True
                                Else
                                    flgAdminBy = False
                                End If
                            Else
                                xWHERE = xWHERE & " AND " & "('" & frmEEBASIC.clpCode(3).Text & "'='" & rsOvtMst("OM_ADMINBY") & "')"
                                If frmEEBASIC.clpCode(3).Text = rsOvtMst("OM_ADMINBY") Then
                                    flgAdminBy = True
                                Else
                                    flgAdminBy = False
                                End If
                            End If
                        Else
                            If xStatusDate Then
                                xWHERE = "('" & rsEE("ED_ADMINBY") & "'='" & rsOvtMst("OM_ADMINBY") & "')"
                                If rsEE("ED_ADMINBY") = rsOvtMst("OM_ADMINBY") Then
                                    flgAdminBy = True
                                Else
                                    flgAdminBy = False
                                End If
                            Else
                                xWHERE = "('" & frmEEBASIC.clpCode(3).Text & "'='" & rsOvtMst("OM_ADMINBY") & "')"
                                If frmEEBASIC.clpCode(3).Text = rsOvtMst("OM_ADMINBY") Then
                                    flgAdminBy = True
                                Else
                                    flgAdminBy = False
                                End If
                            End If
                        End If
                    Else
                        flgAdminBy = True
                    End If
                    If Not IsNull(rsOvtMst("OM_SECTION")) And rsOvtMst("OM_SECTION") <> "" Then
                        If Len(xWHERE) > 0 Then
                            If xStatusDate Then
                                xWHERE = xWHERE & " AND " & "('" & rsEE("ED_SECTION") & "'='" & rsOvtMst("OM_SECTION") & "')"
                                If rsEE("ED_SECTION") = rsOvtMst("OM_SECTION") Then
                                    flgSection = True
                                Else
                                    flgSection = False
                                End If
                            Else
                                xWHERE = xWHERE & " AND " & "('" & frmEEBASIC.clpCode(4).Text & "'='" & rsOvtMst("OM_SECTION") & "')"
                                If frmEEBASIC.clpCode(4).Text = rsOvtMst("OM_SECTION") Then
                                    flgSection = True
                                Else
                                    flgSection = False
                                End If
                            End If
                        Else
                            If xStatusDate Then
                                xWHERE = "('" & rsEE("ED_SECTION") & "'='" & rsOvtMst("OM_SECTION") & "')"
                                If rsEE("ED_SECTION") = rsOvtMst("OM_SECTION") Then
                                    flgSection = True
                                Else
                                    flgSection = False
                                End If
                            Else
                                xWHERE = "('" & frmEEBASIC.clpCode(4).Text & "'='" & rsOvtMst("OM_SECTION") & "')"
                                If frmEEBASIC.clpCode(4).Text = rsOvtMst("OM_SECTION") Then
                                    flgSection = True
                                Else
                                    flgSection = False
                                End If
                            End If
                        End If
                    Else
                        flgSection = True
                    End If
                    If Len(xWHERE) > 0 Then
                        'If xWHERE Then
                        If flgEmp And flgLoc And flgPT And flgRegion And flgSection And flgAdminBy Then
                            Get_Maximum_Bank_Hours = rsOvtMst("OM_MAX_BANK_HRS")
                            rsOvtMst.Close
                            Exit Function
                        End If
                    End If
                
'                    If (IsNull(rsOvtMst("OM_EMP")) Or rsOvtMst("OM_EMP") = "") And (frmEESTATS.clpPT.Text = rsOvtMst("OM_PT")) Then
'                        Get_Maximum_Bank_Hours = rsOvtMst("OM_MAX_BANK_HRS")
'                        rsOvtMst.Close
'                        Exit Function
'                    End If
'
'                    If (IsNull(rsOvtMst("OM_EMP")) Or rsOvtMst("OM_EMP") = "") And (IsNull(rsOvtMst("OM_PT")) Or rsOvtMst("OM_PT") = "") Then
'                        Get_Maximum_Bank_Hours = rsOvtMst("OM_MAX_BANK_HRS")
'                        rsOvtMst.Close
'                        Exit Function
'                    End If
'
'                    If (IsNull(rsOvtMst("OM_PT")) Or rsOvtMst("OM_PT") = "") And (frmEESTATS.clpCode(1).Text = rsOvtMst("OM_EMP")) Then
'                        Get_Maximum_Bank_Hours = rsOvtMst("OM_MAX_BANK_HRS")
'                        rsOvtMst.Close
'                        Exit Function
'                    End If
'
'                    If (frmEESTATS.clpPT.Text = rsOvtMst("OM_PT")) And (frmEESTATS.clpCode(1).Text = rsOvtMst("OM_EMP")) Then
'                        Get_Maximum_Bank_Hours = rsOvtMst("OM_MAX_BANK_HRS")
'                        rsOvtMst.Close
'                        Exit Function
'                    End If
'
'                    If ((IsNull(frmEESTATS.clpPT.Text) Or frmEESTATS.clpPT.Text = "") And (IsNull(rsOvtMst("OM_PT")) Or rsOvtMst("OM_PT") = "")) And ((IsNull(frmEESTATS.clpCode(1).Text) Or frmEESTATS.clpCode(1).Text = "") And (IsNull(rsOvtMst("OM_EMP")) Or rsOvtMst("OM_EMP") = "")) Then
'                        Get_Maximum_Bank_Hours = rsOvtMst("OM_MAX_BANK_HRS")
'                        rsOvtMst.Close
'                        Exit Function
'                    End If

                    rsOvtMst.MoveNext
                Loop
            End If
            rsOvtMst.Close
            Get_Maximum_Bank_Hours = ""
        Else
            Get_Maximum_Bank_Hours = ""
        End If
    End If
    rsEE.Close
End Function

Public Function Get_EmpMaximumOvertime(xEmpNbr)
    Dim SQLQ As String
    Dim rsOvtBank As New ADODB.Recordset
    
    Get_EmpMaximumOvertime = 0
    
    SQLQ = "SELECT OT_EMPNBR, OT_MBANK FROM HR_OVERTIME_BANK WHERE OT_EMPNBR = " & xEmpNbr
    rsOvtBank.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsOvtBank.EOF Then
        Get_EmpMaximumOvertime = IIf(IsNull(rsOvtBank("OT_MBANK")), 0, rsOvtBank("OT_MBANK"))
        rsOvtBank.Update
    End If
    rsOvtBank.Close
    Set rsOvtBank = Nothing
    
End Function

Public Function getCompanyMasterData(xField)
    Dim rsHRPARCO As New ADODB.Recordset
    
    rsHRPARCO.Open "SELECT " & xField & " FROM HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRPARCO.EOF Then
        getCompanyMasterData = IIf(IsNull(rsHRPARCO(xField)), "", rsHRPARCO(xField))
    End If
    rsHRPARCO.Close
    Set rsHRPARCO = Nothing
    
End Function

Public Function GetEmpData(EmpNbr, Field As String, Optional DEFAULT)
    Dim rsEE As New ADODB.Recordset
    rsEE.Open "SELECT " & Field & " FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not IsMissing(DEFAULT) Then
        GetEmpData = DEFAULT
    'Else
    End If
    If rsEE.EOF Then
        'If glbtermopen And glbVadim Then
        If glbVadim Then      'Ticket #19623
            rsEE.Close
            rsEE.Open "SELECT " & Field & " FROM Term_HREMP WHERE ED_EMPNBR=" & EmpNbr & " ORDER BY TERM_SEQ DESC", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            GetEmpData = ""
        End If
    End If
    If Not rsEE.EOF Then
        If Not IsNull(rsEE(Field)) Then GetEmpData = rsEE(Field)
    End If
End Function

Public Function GetTermEmpData(EmpNbr, xTermSEQ, Field As String, Optional DEFAULT)
    Dim rsEE As New ADODB.Recordset
    rsEE.Open "SELECT " & Field & " FROM Term_HREMP WHERE ED_EMPNBR=" & EmpNbr & " AND TERM_SEQ = " & xTermSEQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    
    If Not IsMissing(DEFAULT) Then
        GetTermEmpData = DEFAULT
    'Else
    End If
    If rsEE.EOF Then
        'If glbtermopen And glbVadim Then
        'If glbVadim Then      'Ticket #19623
        '    rsEE.Close
        '    rsEE.Open "SELECT " & Field & " FROM Term_HREMP WHERE ED_EMPNBR=" & EmpNbr & " ORDER BY TERM_SEQ DESC", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        'Else
        '    GetEmpData = ""
        'End If
    End If
    If Not rsEE.EOF Then
        If Not IsNull(rsEE(Field)) Then GetTermEmpData = rsEE(Field)
    End If
End Function

Function GetSuperEmpName(xEmpnoStr)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String, xEmpName As String
Dim xEmpNo
    xEmpNo = getEmpnbr(xEmpnoStr)
    xEmpName = ""
    If Len(xEmpNo) > 0 Then
        SQLQ = "SELECT ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & xEmpNo
        rsEmp.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
        If Not rsEmp.EOF Then
            xEmpName = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
        Else
            rsEmp.Close
            SQLQ = "SELECT ED_SURNAME,ED_FNAME FROM TERM_HREMP WHERE ED_EMPNBR=" & xEmpNo
            rsEmp.Open SQLQ, gdbAdoIhr001X, adOpenForwardOnly
            If Not rsEmp.EOF Then
                xEmpName = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
            End If
        End If
        rsEmp.Close
        If xEmpName = "" Then
            xEmpName = "Unassigned"
        End If
    End If
    GetSuperEmpName = xEmpName
End Function

Function GetEmpData_PayrollID(xPayrollID, Field As String, Optional DEFAULT)
    Dim rsEE As New ADODB.Recordset
    rsEE.Open "SELECT " & Field & " FROM HREMP WHERE ED_PAYROLL_ID='" & xPayrollID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not IsMissing(DEFAULT) Then
        GetEmpData_PayrollID = DEFAULT
    'Else
    End If
    If rsEE.EOF Then
        'If glbtermopen And glbVadim Then
        If glbVadim Then    'Ticket #19623
            rsEE.Close
            rsEE.Open "SELECT " & Field & " FROM Term_HREMP WHERE ED_PAYROLL_ID='" & xPayrollID & "'" & " ORDER BY TERM_SEQ DESC", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            GetEmpData_PayrollID = ""
        End If
    End If
    If Not rsEE.EOF Then
        If Not IsNull(rsEE(Field)) Then GetEmpData_PayrollID = rsEE(Field)
    End If
End Function

Public Sub GetPayID(DATE1, DATE2)
Dim SQLQ, xEmpNbr
Dim rsEmp As New ADODB.Recordset
Dim rsHRAUDIT As New ADODB.Recordset
Dim K, xNum
    SQLQ = "SELECT * FROM HRAUDIT WHERE AU_PAYROLL_ID is null "
    SQLQ = SQLQ & "AND AU_LDATE >=" & Date_SQL(DATE1) & " "
    SQLQ = SQLQ & "AND AU_LDATE <=" & Date_SQL(DATE2) & " "
    rsHRAUDIT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    K = 0
    If Not rsHRAUDIT.EOF Then
        xNum = rsHRAUDIT.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    Do While Not rsHRAUDIT.EOF
        MDIMain.panHelp(0).FloodPercent = (K / xNum) * 100: K = K + 1
        xEmpNbr = rsHRAUDIT("AU_EMPNBR")
        SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNbr & ""
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmp.EOF Then
            If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then
                rsHRAUDIT("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
                rsHRAUDIT.Update
            End If
        Else
            rsEmp.Close
            SQLQ = "SELECT ED_EMPNBR, ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR =" & xEmpNbr & ""
            rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmp.EOF Then
                If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then
                    rsHRAUDIT("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
                    rsHRAUDIT.Update
                End If
            End If
        End If
        rsEmp.Close
    
        rsHRAUDIT.MoveNext
    Loop
    rsHRAUDIT.Close
    MDIMain.panHelp(0).FloodType = 0
End Sub
Public Sub UpdateEMPNBR(nTable As String, nFld As String, nFldTitle, nNewEmpNum, nOldEmpnbr, Optional NDiv)
Dim SQLQ
On Error Resume Next 'Ticket #16462
SQLQ = "UPDATE " & nTable & " SET "
SQLQ = SQLQ & nFld & "=" & getEmpnbr(nNewEmpNum)

If nTable = "HREMP" And glbLinamar Then
    SQLQ = SQLQ & ",ED_DIV = '" & NDiv & "' "
    
    If Left(nOldEmpnbr, 3) <> NDiv Then
        If Not EmpHisCalc(2, getEmpnbr(nNewEmpNum), "", NDiv, "", "", "", "", "", Date) Then MsgBox "EMPHIS Error "
    End If
End If
If nTable <> "HREEO" And nTable <> "ANN_EMP" And nTable <> "ANN_EMP_TEMP" And nTable <> "MTH_SICK" And _
    nTable <> "STATUS_CHANGE" And nTable <> "STATUS_CHANGE_BENEFIT" And nTable <> "TOTALHRS" And _
    nTable <> "TOTALHRS_02" And nTable <> "TOTALHRS_TEMP" And nTable <> "TOTALHRS_TEMP_PaySum" And _
    nTable <> "TOTSICKHRS" Then
    If nFldTitle = "CR_" Or nFldTitle = "CT_" Or nFldTitle = "RC_" Then
        SQLQ = SQLQ & "," & nFldTitle & "LTime = '" & Time$ & "'"
        SQLQ = SQLQ & "," & nFldTitle & "LDate = " & Date_SQL(Date) & ""
    Else
        SQLQ = SQLQ & "," & nFldTitle & "LTIME = '" & Time$ & "'"
        SQLQ = SQLQ & "," & nFldTitle & "LDATE = " & Date_SQL(Date) & ""
    End If
    SQLQ = SQLQ & "," & nFldTitle & "LUSER = '" & glbUserID & "'"
End If
SQLQ = SQLQ & " WHERE "
SQLQ = SQLQ & nFld & "= " & getEmpnbr(nOldEmpnbr)

If InStr(1, nTable, "HRDOC_") <> 0 Then 'Ticket #18349
    gdbAdoIhr001_DOC.Execute SQLQ
Else
    gdbAdoIhr001.Execute SQLQ
End If
End Sub

Public Sub Append_Accrual(xEmpNbr, xType, xEDate, xHours, xAction, Optional xComments)
Dim rsAccrual As New ADODB.Recordset

If Not IsNumeric(xHours) Then Exit Sub

If glbCompSerial = "S/N - 2388W" And xAction = "N" Then 'Ticket #14334
Else
    'If xHours = 0 Then Exit Sub    'Ticket #17924 - Allow zero value update in the Accrual file as it servers as Audit
End If
rsAccrual.Open "SELECT * FROM HR_ACCRUAL WHERE 0=1", gdbAdoIhr001, adOpenStatic, adLockOptimistic
rsAccrual.AddNew
rsAccrual("AC_COMPNO") = "001"
rsAccrual("AC_EMPNBR") = xEmpNbr
rsAccrual("AC_PAYROLL_ID") = getPayrollIDs(xEmpNbr, , True)
rsAccrual("AC_TYPE") = xType
If Len(xEDate) > 0 Then rsAccrual("AC_EDATE") = xEDate
rsAccrual("AC_HRS") = xHours
rsAccrual("AC_ACTION") = xAction
If Not IsMissing(xComments) Then rsAccrual("AC_COMMENTS") = Left(xComments, 4000)
rsAccrual("AC_LUSER") = glbUserID
rsAccrual("AC_LDATE") = Date
rsAccrual("AC_LTIME") = Time$
rsAccrual.Update
rsAccrual.Close
End Sub

Private Function getProcessDate(xEmpNbr)
Dim rsEmp As New ADODB.Recordset
getProcessDate = Date
rsEmp.Open "SELECT ED_UNION FROM HREMP WHERE ED_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenForwardOnly
If Not rsEmp.EOF Then
    If IsDate(rsEmp("ED_UNION")) Then
        If rsEmp("ED_UNION") > Date Then
            getProcessDate = rsEmp("ED_UNION")
        End If
    End If
End If
End Function

Private Function getProcessDate_TERM(xEmpNbr)
Dim rsEmp As New ADODB.Recordset
getProcessDate_TERM = Date
rsEmp.Open "SELECT ED_UNION FROM TERM_HREMP WHERE ED_EMPNBR=" & xEmpNbr, gdbAdoIhr001, adOpenForwardOnly
If Not rsEmp.EOF Then
    If IsDate(rsEmp("ED_UNION")) Then
        If rsEmp("ED_UNION") > Date Then
            getProcessDate_TERM = rsEmp("ED_UNION")
        End If
    End If
End If
End Function

Public Function ScheduleAlreadyExists(xEmpNbr, xEffectiveDate)
    Dim rsHrScheduler As New ADODB.Recordset
    Dim SQLQ As String
    
    ScheduleAlreadyExists = True
    
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND SD_EDATE = " & Date_SQL(xEffectiveDate)
    rsHrScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHrScheduler.EOF Then
        'Schedule already exists
        ScheduleAlreadyExists = True
    Else
        'Schedule does not exists
        ScheduleAlreadyExists = False
    End If
    rsHrScheduler.Close
    Set rsHrScheduler = Nothing
    
End Function

Public Function ScheduleAlreadyExistsFromToDt(xEmpNbr, xFromToDate, Optional xID)
    Dim rsHrScheduler As New ADODB.Recordset
    Dim SQLQ As String
    
    ScheduleAlreadyExistsFromToDt = True
    
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND SD_EDATE <= " & Date_SQL(xFromToDate)
    SQLQ = SQLQ & " AND SD_TDATE >= " & Date_SQL(xFromToDate)
    If Not IsMissing(xID) Then
        SQLQ = SQLQ & " AND SD_ID <> " & xID
    End If
    rsHrScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHrScheduler.EOF Then
        'Schedule already exists
        ScheduleAlreadyExistsFromToDt = True
    Else
        'Schedule does not exists
        ScheduleAlreadyExistsFromToDt = False
    End If
    rsHrScheduler.Close
    Set rsHrScheduler = Nothing
    
End Function

Public Function LaterScheduleExists(xEmpNbr, xEffectiveDate)
    Dim rsHrScheduler As New ADODB.Recordset
    Dim SQLQ As String
    
    LaterScheduleExists = True
    
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND SD_EDATE > " & Date_SQL(xEffectiveDate)
    rsHrScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHrScheduler.EOF Then
        'Later Schedule already exists
        LaterScheduleExists = True
    Else
        'Later Schedule does not exists
        LaterScheduleExists = False
    End If
    rsHrScheduler.Close
    Set rsHrScheduler = Nothing

End Function

Public Function LaterScheduleExistsFromToDt(xEmpNbr, xFromToDate)
    Dim rsHrScheduler As New ADODB.Recordset
    Dim SQLQ As String
    
    LaterScheduleExistsFromToDt = True
    
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND (SD_EDATE >= " & Date_SQL(xFromToDate)
    SQLQ = SQLQ & " OR SD_TDATE >= " & Date_SQL(xFromToDate) & ")"
    rsHrScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHrScheduler.EOF Then
        'Later Schedule already exists
        LaterScheduleExistsFromToDt = True
    Else
        'Later Schedule does not exists
        LaterScheduleExistsFromToDt = False
    End If
    rsHrScheduler.Close
    Set rsHrScheduler = Nothing

End Function

Public Function OverlapScheduleExists(xEmpNbr, xFromDate, xToDate, Optional xID)
    Dim rsHrScheduler As New ADODB.Recordset
    Dim SQLQ As String
    
    OverlapScheduleExists = True
    
    SQLQ = "SELECT * FROM HR_SCHEDULER WHERE SD_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND SD_TDATE < " & Date_SQL(xToDate)
    SQLQ = SQLQ & " AND SD_TDATE > " & Date_SQL(xFromDate)
    If Not IsMissing(xID) Then
        SQLQ = SQLQ & " AND SD_ID <> " & xID
    End If
    rsHrScheduler.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHrScheduler.EOF Then
        'Work Schedule overlapping
        OverlapScheduleExists = True
    Else
        'Work Schedule overlap does not exists
        OverlapScheduleExists = False
    End If
    rsHrScheduler.Close
    Set rsHrScheduler = Nothing

End Function

Public Function UnapprovedRequestExists(xEmpNbr)
    Dim rsVacReq As New ADODB.Recordset
    Dim SQLQ As String
    
    UnapprovedRequestExists = True
    
    'Create a query to give a list of unapproved requests and with date range greater than 1 day
    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ WHERE VT_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND (VT_DELFLAG=0 OR VT_DELFLAG IS NULL)"
    SQLQ = SQLQ & " AND VT_FROM <> VT_TO"   'date range has more than 1 day
    'Added 'SUBMITTED' as I noticed in the table that when 'SUBMITTED', the VT_APPROVED is not longer blank
    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('SUBMITTED','RESUBMITTED','APP/FWD','REJECTED'))"
    'SQLQ = SQLQ & " AND VT_VACTIME=1 "  'vacation time only
    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsVacReq.EOF Then
        UnapprovedRequestExists = True
    Else
        UnapprovedRequestExists = False
    End If
    rsVacReq.Close
    Set rsVacReq = Nothing
    
    'If a list of unapproved requests are required then show the following fields:
    'Number, Name,From Date, To Date, Requested Date, Processed Date, Supervisor Name, Status, Hours Requested
    'VT_EMPNBR,VT_EMPNAME,VT_FROM,VT_TO,VT_REQDATE,VT_PROCDATE,VT_SUPERNAME,VT_APPROVED(Status),VT_HRS(Hours Requested)?
End Function

Public Function UnapprovedRequestExistsFromToDt(xEmpNbr, xFromToDate, Optional xToDate)
    Dim rsVacReq As New ADODB.Recordset
    Dim SQLQ As String
    
    UnapprovedRequestExistsFromToDt = True
    
    'Create a query to give a list of unapproved requests and with date range greater than 1 day
    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ WHERE VT_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND (VT_DELFLAG=0 OR VT_DELFLAG IS NULL)"
    'SQLQ = SQLQ & " AND VT_FROM <> VT_TO"   'date range has more than 1 day
    'Ticket #22221 - If Vac/Time Req From Date and To Date between the Work Schedule Rule Date then do not allow delete
    If Not IsMissing(xToDate) Then
        SQLQ = SQLQ & " AND (VT_FROM <=  " & Date_SQL(xToDate) & " AND VT_TO <= " & Date_SQL(xToDate) & ")"
    Else
        SQLQ = SQLQ & " AND (VT_FROM >=  " & Date_SQL(xFromToDate) & " AND VT_TO <= " & Date_SQL(xFromToDate) & ")"
    End If
    
    'Added 'SUBMITTED' as I noticed in the table that when 'SUBMITTED', the VT_APPROVED is not longer blank
    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('SUBMITTED','RESUBMITTED','APP/FWD','REJECTED'))"
    'SQLQ = SQLQ & " AND VT_VACTIME=1 "  'vacation time only
    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsVacReq.EOF Then
        UnapprovedRequestExistsFromToDt = True
    Else
        UnapprovedRequestExistsFromToDt = False
    End If
    rsVacReq.Close
    Set rsVacReq = Nothing
    
    'If a list of unapproved requests are required then show the following fields:
    'Number, Name,From Date, To Date, Requested Date, Processed Date, Supervisor Name, Status, Hours Requested
    'VT_EMPNBR,VT_EMPNAME,VT_FROM,VT_TO,VT_REQDATE,VT_PROCDATE,VT_SUPERNAME,VT_APPROVED(Status),VT_HRS(Hours Requested)?
End Function

Public Function UnapprovedRequestExistsChangeDate(xEmpNbr, xChangeDate)
    Dim rsVacReq As New ADODB.Recordset
    Dim SQLQ As String
    
    UnapprovedRequestExistsChangeDate = True
    
    'Create a query to give a list of unapproved requests from the Change Date
    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ WHERE VT_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND (VT_DELFLAG=0 OR VT_DELFLAG IS NULL)"
    SQLQ = SQLQ & " AND ((VT_FROM >=  " & Date_SQL(xChangeDate) & ") OR (VT_TO >= " & Date_SQL(xChangeDate) & "))"
    'Added 'SUBMITTED' as I noticed in the table that when 'SUBMITTED', the VT_APPROVED is not longer blank
    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('SUBMITTED','RESUBMITTED','APP/FWD','REJECTED'))"
    'SQLQ = SQLQ & " AND VT_VACTIME=1 "  'vacation time only
    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsVacReq.EOF Then
        UnapprovedRequestExistsChangeDate = True
    Else
        UnapprovedRequestExistsChangeDate = False
    End If
    rsVacReq.Close
    Set rsVacReq = Nothing
    
    'If a list of unapproved requests are required then show the following fields:
    'Number, Name,From Date, To Date, Requested Date, Processed Date, Supervisor Name, Status, Hours Requested
    'VT_EMPNBR,VT_EMPNAME,VT_FROM,VT_TO,VT_REQDATE,VT_PROCDATE,VT_SUPERNAME,VT_APPROVED(Status),VT_HRS(Hours Requested)?
End Function

Public Function AnyUnapprovedRequestExists(xFromToDate, Optional xToDate)
    Dim rsVacReq As New ADODB.Recordset
    Dim SQLQ As String
    
    AnyUnapprovedRequestExists = True
    
    'Create a query to give a list of unapproved requests that is exists in future
    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ " 'WHERE VT_EMPNBR = " & xEMPNBR
    SQLQ = SQLQ & " WHERE (VT_DELFLAG=0 OR VT_DELFLAG IS NULL)"
    'SQLQ = SQLQ & " AND VT_FROM <> VT_TO"   'date range has more than 1 day
    'Ticket #22221 - If Vac/Time Req From Date and To Date between the Work Schedule Rule Date then do not allow delete
    'If Not IsMissing(xToDate) Then
    '    SQLQ = SQLQ & " AND (VT_FROM <=  " & Date_SQL(xToDate) & " AND VT_TO <= " & Date_SQL(xToDate) & ")"
    'Else
    '    SQLQ = SQLQ & " AND (VT_FROM >=  " & Date_SQL(xFromToDate) & " AND VT_TO <= " & Date_SQL(xFromToDate) & ")"
    'End If
    SQLQ = SQLQ & " AND (VT_TO >=  " & Date_SQL(xFromToDate) & ")"
    
    'Added 'SUBMITTED' as I noticed in the table that when 'SUBMITTED', the VT_APPROVED is not longer blank
    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('SUBMITTED','RESUBMITTED','APP/FWD','REJECTED'))"
    'SQLQ = SQLQ & " AND VT_VACTIME=1 "  'vacation time only
    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsVacReq.EOF Then
        AnyUnapprovedRequestExists = True
    Else
        AnyUnapprovedRequestExists = False
    End If
    rsVacReq.Close
    Set rsVacReq = Nothing
    
End Function

Public Function LastDateUnapprovedRequest(xEmpNbr, xFromToDate, xToDate)
    Dim rsVacReq As New ADODB.Recordset
    Dim SQLQ As String
    
    LastDateUnapprovedRequest = ""
    
    'Create a query to give you the Last Date of the Unapproved / rejected requests
    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ WHERE VT_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND (VT_DELFLAG=0 OR VT_DELFLAG IS NULL)"
    
    'Ticket #22221 - If Vac/Time Req From Date and To Date between the Work Schedule Rule Dates
    SQLQ = SQLQ & " AND (VT_FROM >=  " & Date_SQL(xFromToDate) & " AND VT_TO <= " & Date_SQL(xToDate) & ")"
    
    'Added 'APPROVED' as well as for the purpose I am using this function - Rebuild the WS Deatails table, I don't
    'want to delete the existing WS Detail records for the time period the Request has been approved and it used
    'the respective WS Detail records to create the request.
    'Added 'SUBMITTED' as I noticed in the table that when 'SUBMITTED', the VT_APPROVED is not longer blank
    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('APPROVED','SUBMITTED','RESUBMITTED','APP/FWD','REJECTED'))"
    SQLQ = SQLQ & " ORDER BY VT_TO"
    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsVacReq.EOF Then
        rsVacReq.MoveLast
        
        LastDateUnapprovedRequest = rsVacReq("VT_TO")
    Else
        LastDateUnapprovedRequest = ""
    End If
    rsVacReq.Close
    Set rsVacReq = Nothing
    
    'If a list of unapproved requests are required then show the following fields:
    'Number, Name,From Date, To Date, Requested Date, Processed Date, Supervisor Name, Status, Hours Requested
    'VT_EMPNBR,VT_EMPNAME,VT_FROM,VT_TO,VT_REQDATE,VT_PROCDATE,VT_SUPERNAME,VT_APPROVED(Status),VT_HRS(Hours Requested)?
End Function

Public Function WorkScheduleExists(xDate) As Boolean
    Dim rsSchedule As New ADODB.Recordset
    Dim SQLQ As String
    
    WorkScheduleExists = False
    
    'Any schedule exists with To Date later than the Effective Date?
    SQLQ = "SELECT * FROM HR_SCHEDULER "
    SQLQ = SQLQ & " WHERE SD_TDATE >= " & Date_SQL(xDate)
    rsSchedule.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsSchedule.EOF Then
        WorkScheduleExists = True
    Else
        WorkScheduleExists = False
    End If
    rsSchedule.Close
    Set rsSchedule = Nothing

End Function

'Public Function UnapprovedRequestExistsFromToDt(xEmpnbr, xFromDate, xToDate)
'    Dim rsVacReq As New ADODB.Recordset
'    Dim SQLQ As String
'
'    UnapprovedRequestExistsFromToDt = True
'
'    'Create a query to give a list of unapproved requests and with date range greater than 1 day
'    SQLQ = "SELECT * FROM HR_VACTIMEOFF_REQ WHERE VT_EMPNBR = " & xEmpnbr
'    SQLQ = SQLQ & " AND VT_DELFLAG=0"
'    'SQLQ = SQLQ & " AND VT_FROM <> VT_TO"   'date range has more than 1 day
'    'Ticket #22221 - If Vac/Time Req From Date and To Date between the Work Schedule Rule Date then do not allow delete
'    SQLQ = SQLQ & " AND ((VT_FROM >=  " & Date_SQL(xFromDate) & " AND VT_FROM <=  " & Date_SQL(xToDate) & ") OR (VT_TO >= " & Date_SQL(xFromDate) & " AND VT_TO <= " & Date_SQL(xToDate) & "))"
'    SQLQ = SQLQ & " AND (VT_APPROVED IS NULL OR VT_APPROVED = '' OR VT_APPROVED IN ('RESUBMITTED','APP/FWD','REJECTED'))"
'    'SQLQ = SQLQ & " AND VT_VACTIME=1 "  'vacation time only
'    rsVacReq.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'    If Not rsVacReq.EOF Then
'        UnapprovedRequestExistsFromToDt = True
'    Else
'        UnapprovedRequestExistsFromToDt = False
'    End If
'    rsVacReq.Close
'    Set rsVacReq = Nothing
'
'    'If a list of unapproved requests are required then show the following fields:
'    'Number, Name,From Date, To Date, Requested Date, Processed Date, Supervisor Name, Status, Hours Requested
'    'VT_EMPNBR,VT_EMPNAME,VT_FROM,VT_TO,VT_REQDATE,VT_PROCDATE,VT_SUPERNAME,VT_APPROVED(Status),VT_HRS(Hours Requested)?
'End Function

Public Function isWorkSchedule(xEmpNo) As Boolean
    Dim rsHRTABL As New ADODB.Recordset
    Dim SQLQ As String
    
    isWorkSchedule = False
    SQLQ = "SELECT TB_WORKSCHED FROM HRTABL WHERE TB_NAME = 'EDEM'"
    SQLQ = SQLQ & " AND TB_KEY = (SELECT ED_EMP FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & ")"
    rsHRTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsHRTABL.EOF Then
        If Not IsNull(rsHRTABL("TB_WORKSCHED")) Then
            If rsHRTABL("TB_WORKSCHED") Then
                isWorkSchedule = True
            Else
                isWorkSchedule = False
            End If
        Else
            isWorkSchedule = False
        End If
    Else
        isWorkSchedule = False
    End If
    rsHRTABL.Close
    Set rsHRTABL = Nothing
End Function

Public Function Get_Payroll_ID_For_Benefit(xEmpNo) As String
    Dim SQLQ As String
    Dim rsJobHist As New ADODB.Recordset
    Dim rsBenefit As New ADODB.Recordset
    Dim rsCurJobRec As New ADODB.Recordset
    Dim xPayrollID As String
    Dim xRecCnt As Integer
    
    'Ticket #19687
    'Retrieve current Job record(s) of the employee and check for 'For Benefit'.
    'If one Current record only, then update employee's Benefit record with this Job's Payroll ID
    'If more then one Current record, then check for position with 'For Benefit' checked. If found then update
    'employee's Benefit record with that Job's Payroll ID.
    
    Get_Payroll_ID_For_Benefit = ""
    xPayrollID = ""
    
    'Get the count of total current records
    SQLQ = "SELECT COUNT(JH_EMPNBR) AS TOT_REC FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & xEmpNo
    rsCurJobRec.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsCurJobRec.EOF Then
        xRecCnt = rsCurJobRec("TOT_REC")
    Else
        'No current records found, get the Payroll ID from Employee Demographics screen
        xPayrollID = GetEmpData(xEmpNo, "ED_PAYROLL_ID", "")
    End If
    rsCurJobRec.Close
    Set rsCurJobRec = Nothing
    
    'Retrieve employee's current Job records
    SQLQ = "SELECT JH_EMPNBR, JH_PAYROLL_ID, JH_USRCHECK FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_CURRENT <> 0 AND JH_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " ORDER BY JH_SDATE DESC"
    rsJobHist.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    Do While Not rsJobHist.EOF
        If xRecCnt = 1 Then
            'Only one current record, get the Payroll ID of that record
            xPayrollID = rsJobHist("JH_PAYROLL_ID")
        Else
            'Multiple current Job records, get the Payroll ID of the 'For Benefit' job checked
            If rsJobHist("JH_USRCHECK") <> 0 Then
                xPayrollID = rsJobHist("JH_PAYROLL_ID")
                Exit Do
            End If
        End If
        rsJobHist.MoveNext
    Loop
    rsJobHist.Close
    Set rsJobHist = Nothing
    
    'No Payroll ID found, get it from Employee's Demographics screen
    If xPayrollID = "" Then
        xPayrollID = GetEmpData(xEmpNo, "ED_PAYROLL_ID", "")
    End If
    
    Get_Payroll_ID_For_Benefit = xPayrollID
    
End Function

Public Function get_BasicTaxAmounts(xTaxField)
    Dim SQLQ As String
    Dim rsPA As New ADODB.Recordset
    
    get_BasicTaxAmounts = ""
    
    'PC_FEDTAX,PC_PROVTAX
    SQLQ = "SELECT " & xTaxField & " FROM HRPARCO"
    rsPA.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    If Not rsPA.EOF Then
        get_BasicTaxAmounts = rsPA(xTaxField)
    End If
    rsPA.Close
    Set rsPA = Nothing
    
End Function

Public Sub CreateTableMasterCode(xTblName As String, xTblKey As String, xTblKeyDesc As String)
    Dim SQLQ As String
    Dim rsTABL As New ADODB.Recordset
    
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & xTblName & "' AND TB_KEY = '" & xTblKey & "'"
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsTABL.EOF Then
        rsTABL.AddNew
        rsTABL("TB_COMPNO") = "001"
        rsTABL("TB_NAME") = xTblName
        rsTABL("TB_KEY") = xTblKey
        rsTABL("TB_DESC") = xTblKeyDesc
        rsTABL("TB_LDATE") = Date
        rsTABL("TB_LTIME") = Time$
        rsTABL("TB_LUSER") = glbUserID
        rsTABL.Update
    End If
    rsTABL.Close
    Set rsTABL = Nothing
    
End Sub

Public Sub EmployeeFlagUpd(xEmpNo, xFlagNo, xVal, xDate, xFUDate, xFUReason, Optional xTERM_Seq)
    Dim rsEmpFlag As New ADODB.Recordset
    Dim SQLQ As String
    
    If Not IsMissing(xTERM_Seq) Then
        SQLQ = "SELECT * FROM TERM_HREMP_FLAGS WHERE EF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_Seq & " "
    Else
        SQLQ = "SELECT * FROM HREMP_FLAGS WHERE EF_EMPNBR = " & xEmpNo & " "
    End If
    rsEmpFlag.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsEmpFlag.EOF Then
        rsEmpFlag.AddNew
        rsEmpFlag("EF_COMPNO") = "001"
        rsEmpFlag("EF_EMPNBR") = xEmpNo
        If Not IsMissing(xTERM_Seq) Then
            rsEmpFlag("TERM_SEQ") = xTERM_Seq
        End If
    End If
    
    'Flag Value
    If Len(xVal) > 0 Then
        rsEmpFlag("EF_FLAGVAL" & xFlagNo) = xVal
    End If
    
    'Flag Date
    If Len(xDate) > 0 Then
        If IsDate(xDate) Then
            rsEmpFlag("EF_FLAGDTE" & xFlagNo) = CVDate(xDate)
        End If
    End If
    
    'Follow Up Date & Reason
    If Len(xFUDate) > 0 Then
        If IsDate(xFUDate) Then
            rsEmpFlag("EF_FUDTE" & xFlagNo) = CVDate(xFUDate)
            'Ticket #23641 - Field name issue
            If xFlagNo = 19 Then
                rsEmpFlag("EF_FTREAS" & xFlagNo) = xFUReason
            Else
                rsEmpFlag("EF_FUREAS" & xFlagNo) = xFUReason
            End If
        End If
    End If
    rsEmpFlag("EF_LDATE") = Date
    rsEmpFlag("EF_LTIME") = Time$
    rsEmpFlag("EF_LUSER") = glbUserID
    rsEmpFlag.Update
    rsEmpFlag.Close

End Sub

Public Function OMER_UseCostTable()
Dim rsOMERRule As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    SQLQ = "SELECT * FROM HR_OMERS_FORMULA WHERE OM_YEAR = " & Year(Date)
    rsOMERRule.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsOMERRule.EOF Then
        retVal = False
    Else
        retVal = True
    End If
    OMER_UseCostTable = retVal
End Function

Public Function EmpOmersCalculate(xEmpNo, xCode, xUptBene, Optional xAnnSalary, Optional xBF_BENE_ID)
Dim rsOMERRule As New ADODB.Recordset
Dim rsEBene As New ADODB.Recordset
Dim SQLQ As String
Dim xTier1, xTier2, xTier3
Dim retVal
    'see logic in OMERS Calculation in Benefit Group Master.docx
    'in X:\Word Documents\INFOHR Documentation folder
    SQLQ = "SELECT * FROM HR_OMERS_FORMULA WHERE OM_YEAR = " & Year(Date)
    rsOMERRule.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsOMERRule.EOF Then
        rsOMERRule.Close
        retVal = 0
        GoTo Fun_end
    End If
    
    If IsMissing(xAnnSalary) Then
        xAnnSalary = CrtSalary(xEmpNo)
    End If
    'Tier 1: YMPE
    xTier1 = rsOMERRule("OM_YMPE_MAX") * rsOMERRule("OM_PERCENT_YMPE")

    'Tier 3: OM_TIER3_PERC
    xTier3 = 0
    If xAnnSalary > rsOMERRule("OM_MAXREG_RRP") Then
        'Tier 2: RRP
        xTier2 = (rsOMERRule("OM_MAXREG_RRP") - rsOMERRule("OM_YMPE_MAX")) * rsOMERRule("OM_PERC_MAX_RRP")
        If xTier2 < 0 Then xTier2 = 0
        xTier3 = (xAnnSalary - rsOMERRule("OM_MAXREG_RRP")) * rsOMERRule("OM_TIER3_PERC")
    Else
        'Tier 2: RRP
        xTier2 = (xAnnSalary - rsOMERRule("OM_YMPE_MAX")) * rsOMERRule("OM_PERC_MAX_RRP")
        If xTier2 < 0 Then xTier2 = 0
    End If
    
    retVal = xTier1 + xTier2 + xTier3
    
    If xUptBene = "Y" Then 'Update Total Cost for OMER benefit
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_BCODE = '" & xCode & "' "
        If Not IsMissing(xBF_BENE_ID) Then
            SQLQ = SQLQ & "AND BF_BENE_ID = " & xBF_BENE_ID & " "
        End If
        rsEBene.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEBene.EOF Then
            rsEBene("BF_TCOST") = retVal
            If IsNumeric(rsEBene("BF_PCC")) Then
                rsEBene("BF_CCOST") = retVal * rsEBene("BF_PCC")
                rsEBene("BF_MTHCCOST") = (retVal * rsEBene("BF_PCC")) / 12
            End If
            If IsNumeric(rsEBene("BF_PCE")) Then
                rsEBene("BF_ECOST") = retVal * rsEBene("BF_PCE")
                rsEBene("BF_MTHECOST") = (retVal * rsEBene("BF_PCE")) / 12
            End If
            rsEBene.Update
        End If
        rsEBene.Close
    End If
Fun_end:
    EmpOmersCalculate = retVal
End Function

Public Function Round_HR(xVal, xDec, Optional xType)
Dim retVal
Dim xTmp
Dim xVa2
    If UCase(xType) = "T" Then 'Truncate, e.g. Round_HR(0.666667,2, "T") = 0.66
        xTmp = "1" & String(xDec, "0")
        xTmp = Val(xTmp)
        xVa2 = xVal * xTmp
        xVa2 = Int(xVa2)
        retVal = xVa2 / xTmp
    End If
    
    'to fix the VB Round() problem, Round(0.615,2) = 0.61. it should be 0.62
    If UCase(xType) = "D" Then 'Round, e.g. Round_HR(0.6152,2, "D") = 0.62
        xTmp = "1" & String(xDec, "0")
        xTmp = Val(xTmp)
        xVa2 = xVal * xTmp
        If (xVa2 - Int(xVa2)) >= 0.5 Then
            xVa2 = Int(xVa2) + 1
        Else
            xVa2 = Int(xVa2)
        End If
        retVal = xVa2 / xTmp
    End If
    Round_HR = retVal
End Function

Public Function IsLOATypeCode(xCode) As Boolean
    Dim SQLQ As String
    Dim rsTA As New ADODB.Recordset
    
    SQLQ = "SELECT TB_USR3 FROM HRTABL WHERE TB_NAME='EDEM' AND TB_KEY = '" & xCode & "'"
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsTA.EOF Then
        If rsTA("TB_USR3") = True Then
            IsLOATypeCode = True
        Else
            IsLOATypeCode = False
        End If
    Else
        IsLOATypeCode = False
    End If
    rsTA.Close
    Set rsTA = Nothing

End Function

Public Sub SpecificFunction_Template_Based_Security_Profile_Update(xUserID, xTemplate, xUpdType, Optional xFunction)
Dim x, sUSERID
Dim SQLQ As String
Dim rsSecAccess As New ADODB.Recordset
Dim rsINSERT As New ADODB.Recordset
Dim rsSecTemplate As New ADODB.Recordset

    'Ticket #20585 - Security Based on Template Profile
    
    If Not IsMissing(xFunction) Then
        
        If xFunction = "CODES" Then
            'Retrieve Template Security Profile from HR_SECURE_ACCESS
            SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            SQLQ = SQLQ & " AND CODENAME IS NOT NULL"
            rsSecTemplate.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            If rsSecTemplate.EOF Then
                MsgBox "This Template has no Security Profile setup.", vbOKOnly, "Error finding Security Profile"
                Exit Sub
            Else
                'Delete User's Profile first and then add back based on Template Profile
                SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                SQLQ = SQLQ & " AND CODENAME IS NOT NULL"
                gdbAdoIhr001.Execute SQLQ
                
                'Open User's Security record to add back
                SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                SQLQ = SQLQ & " AND CODENAME IS NOT NULL"
                rsSecAccess.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
            End If
            
            MDIMain.panHelp(0).Caption = "Please wait while system updates security profile..."
            
            'Add the user back using the template profile
            Do While Not rsSecTemplate.EOF
                 rsSecAccess.AddNew
                 rsSecAccess("USERID") = xUserID
                 rsSecAccess("FUNCTION") = rsSecTemplate("FUNCTION")
                 rsSecAccess("ACCESSABLE") = rsSecTemplate("ACCESSABLE")
                 rsSecAccess("Maintainable") = rsSecTemplate("Maintainable")
                 rsSecAccess("CODENAME") = rsSecTemplate("CODENAME")
                 rsSecAccess("LDATE") = Date
                 rsSecAccess("LTIME") = Time$
                 rsSecAccess("LUSER") = glbUserID
        
                 rsSecAccess.Update
                 rsSecTemplate.MoveNext
            Loop
            rsSecAccess.Close
            Set rsSecAccess = Nothing
            rsSecTemplate.Close
            Set rsSecTemplate = Nothing
        End If
        
        If xFunction = "DEPARTMENT" Then
            'Add the Department Security
            Dim rsFrmSecDept As New ADODB.Recordset
            Dim rsToSecDept As New ADODB.Recordset
            
            'Retrieve Template's Department Security
            SQLQ = "SELECT * FROM HRPASDEP WHERE PD_USERID='" & Replace(xTemplate, "'", "''") & "'"
            rsFrmSecDept.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            'Delete User's Department Security first and then add back based on Template
            SQLQ = "DELETE FROM HRPASDEP WHERE PD_USERID='" & Replace(xUserID, "'", "''") & "'"
            gdbAdoIhr001.Execute SQLQ
            
            'Open User's Department Security record
            SQLQ = "SELECT * FROM HRPASDEP WHERE PD_USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecDept.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
            
            'Add User's Dept Security based on the Template
            Do While Not rsFrmSecDept.EOF
                rsToSecDept.AddNew
                rsToSecDept("PD_COMPNO") = "001"
                rsToSecDept("PD_USERID") = xUserID
                rsToSecDept("PD_DEPT") = rsFrmSecDept("PD_DEPT")
                rsToSecDept("PD_ORG") = rsFrmSecDept("PD_ORG")
                rsToSecDept("PD_DIV") = rsFrmSecDept("PD_DIV")
                rsToSecDept("PD_SECTION") = rsFrmSecDept("PD_SECTION")
                rsToSecDept("PD_ADMINBY") = rsFrmSecDept("PD_ADMINBY")
                'Ticket #22682 - Release 8.0
                rsToSecDept("PD_LOC") = rsFrmSecDept("PD_LOC")
                rsToSecDept("PD_REGION") = rsFrmSecDept("PD_REGION")
                
                'Ticket #24161 - Samuel Only - Release 8.0
                If glbSamuel Then
                    rsToSecDept("PD_SUPCODE") = rsFrmSecDept("PD_SUPCODE")
                    rsToSecDept("PD_VADIM2") = rsFrmSecDept("PD_VADIM2")
                End If
                
                rsToSecDept("PD_INCLEMPNBR") = rsFrmSecDept("PD_INCLEMPNBR")
                rsToSecDept("PD_EXCLEMPNBR") = rsFrmSecDept("PD_EXCLEMPNBR")
                rsToSecDept.Update
                
                rsFrmSecDept.MoveNext
            Loop
            rsToSecDept.Close
            Set rsToSecDept = Nothing
            rsFrmSecDept.Close
            Set rsFrmSecDept = Nothing
        End If
        
        If xFunction = "COMMENTS" Then
            'Add the Comments Security
            Dim rsFrmSecComments As New ADODB.Recordset
            Dim rsToSecComments As New ADODB.Recordset
            
            'Retrieve Template's Comments Security
            SQLQ = "SELECT * FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            rsFrmSecComments.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            'Delete User's Comments Security first and then add back based on Template
            SQLQ = "DELETE FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            gdbAdoIhr001.Execute SQLQ
            
            'Open User's Department Security record
            SQLQ = "SELECT * FROM HR_SECURE_COMMENTS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecComments.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                
            'Add User's Comments Security based on the Template
            Do While Not rsFrmSecComments.EOF
                rsToSecComments.AddNew
                rsToSecComments("COMPNO") = "001"
                rsToSecComments("USERID") = xUserID
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
            Set rsToSecComments = Nothing
            rsFrmSecComments.Close
            Set rsFrmSecComments = Nothing
        End If
        
        If xFunction = "CUSTOMRPTS" Then
            'Add the Custom Reports Security
            Dim rsFrmSecCustmRpt As New ADODB.Recordset
            Dim rsToSecCustmRpt As New ADODB.Recordset
            
            'Retrieve Template's Custom Reports Security
            SQLQ = "SELECT * FROM HR_SECRPT WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            rsFrmSecCustmRpt.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            'Delete User's Custom Reports Security first and then add back based on Template
            SQLQ = "DELETE FROM HR_SECRPT WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            gdbAdoIhr001.Execute SQLQ
            
            'Open User's Custom Reports Security record
            SQLQ = "SELECT * FROM HR_SECRPT WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecCustmRpt.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                
            'Add User's Custom Reports Security based on the Template
            Do While Not rsFrmSecCustmRpt.EOF
                rsToSecCustmRpt.AddNew
                rsToSecCustmRpt("COMPNO") = "001"
                rsToSecCustmRpt("USERID") = xUserID
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
            Set rsToSecCustmRpt = Nothing
            rsFrmSecCustmRpt.Close
            Set rsFrmSecCustmRpt = Nothing
        End If
        
        If xFunction = "CUSTOMFEATURE" Or xFunction = "CUSTOMFEATURE_PEN" Then
            Dim rsFrmSecCustmFeat As New ADODB.Recordset
            Dim rsToSecCustmFeat As New ADODB.Recordset

            'Add Linamar's Custom Features Security
            If glbLinamar Then
                'Retrieve Template's Linamar's Custom Features Security
                SQLQ = "SELECT * FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
                rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                
                'Delete User's Linamar's Custom Features Security first and then add back based on Template
                SQLQ = "DELETE FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                gdbAdoIhr001.Execute SQLQ
                
                'Open User's Linamar's Custom Features Security record
                SQLQ = "SELECT * FROM LN_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                        
                'Add User's Linamar's Custom Feature Security based on the Template
                Do While Not rsFrmSecCustmFeat.EOF
                    rsToSecCustmFeat.AddNew
                    rsToSecCustmFeat("COMPNO") = "001"
                    rsToSecCustmFeat("USERID") = xUserID
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
                Set rsToSecCustmFeat = Nothing
                rsFrmSecCustmFeat.Close
                Set rsFrmSecCustmFeat = Nothing
            End If
            
        
            'Add WHSCC's Custom Features Security
            If glbWHSCC Then
                'Retrieve Template's WHSCC's Custom Features Security
                SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
                SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WHSC'"
                rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                
                'Delete User's WHSCC's Custom Features Security first and then add back based on Template
                SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WHSC'"
                gdbAdoIhr001.Execute SQLQ
                
                'Open User's WHSCC's Custom Features Security record
                SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WHSC'"
                rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                        
                'Add User's WHSCC's Custom Feature Security based on the Template
                Do While Not rsFrmSecCustmFeat.EOF
                    rsToSecCustmFeat.AddNew
                    rsToSecCustmFeat("COMPNO") = "001"
                    rsToSecCustmFeat("USERID") = xUserID
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
                Set rsToSecCustmFeat = Nothing
                rsFrmSecCustmFeat.Close
                Set rsFrmSecCustmFeat = Nothing
            End If
            
            
            'Add WFC's Custom Features Security
            If glbWFC Then
                'Retrieve Template's WFC's Custom Features Security
                SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
                If xFunction = "CUSTOMFEATURE_PEN" Then
                    SQLQ = SQLQ & " AND LEFT([FUNCTION],7)='WFCPEN_'"
                Else
                    SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WFC_'"
                End If
                rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                
                'Delete User's WFC's Custom Features Security first and then add back based on Template
                SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                If xFunction = "CUSTOMFEATURE_PEN" Then
                    SQLQ = SQLQ & " AND LEFT([FUNCTION],7)='WFCPEN_'"
                Else
                    SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WFC_'"
                End If
                gdbAdoIhr001.Execute SQLQ
                
                'Open User's WFC's Custom Features Security record
                SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                If xFunction = "CUSTOMFEATURE_PEN" Then
                    SQLQ = SQLQ & " AND LEFT([FUNCTION],7)='WFCPEN_'"
                Else
                    SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='WFC_'"
                End If
                rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                        
                'Add User's WFC's Custom Feature Security based on the Template
                Do While Not rsFrmSecCustmFeat.EOF
                    rsToSecCustmFeat.AddNew
                    rsToSecCustmFeat("COMPNO") = "001"
                    rsToSecCustmFeat("USERID") = xUserID
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
                Set rsToSecCustmFeat = Nothing
                rsFrmSecCustmFeat.Close
                Set rsFrmSecCustmFeat = Nothing
            End If
            
            
            'Add Samuel's Custom Features Security
            If glbCompSerial = "S/N - 2382W" Then
                'Retrieve Template's Samuel's Custom Features Security
                SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
                SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='SAM_'"
                rsFrmSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                
                'Delete User's Samuel's Custom Features Security first and then add back based on Template
                SQLQ = "DELETE FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='SAM_'"
                gdbAdoIhr001.Execute SQLQ
                
                'Open User's Samuel's Custom Features Security record
                SQLQ = "SELECT * FROM HR_SECURE_ACCESS WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
                SQLQ = SQLQ & " AND LEFT([FUNCTION],4)='SAM_'"
                rsToSecCustmFeat.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                        
                'Add User's Samuel's Custom Feature Security based on the Template
                Do While Not rsFrmSecCustmFeat.EOF
                    rsToSecCustmFeat.AddNew
                    rsToSecCustmFeat("COMPNO") = "001"
                    rsToSecCustmFeat("USERID") = xUserID
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
                Set rsToSecCustmFeat = Nothing
                rsFrmSecCustmFeat.Close
                Set rsFrmSecCustmFeat = Nothing
            End If
            
        End If
        
        If xFunction = "FOLLOWUP" Then
            'Add Follow Up Security
            Dim rsFrmSecFollowUp As New ADODB.Recordset
            Dim rsToSecFollowUp As New ADODB.Recordset
            
            'Retrieve Template's Follow Up Security
            SQLQ = "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            rsFrmSecFollowUp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            'Delete User's Follow Up Security first and then add back based on Template
            SQLQ = "DELETE FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            gdbAdoIhr001.Execute SQLQ
            
            'Open User's Follow Up Security record
            SQLQ = "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecFollowUp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                
            'Add User's Follow Up Security based on the Template
            Do While Not rsFrmSecFollowUp.EOF
                rsToSecFollowUp.AddNew
                rsToSecFollowUp("COMPNO") = "001"
                rsToSecFollowUp("USERID") = xUserID
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
            Set rsToSecFollowUp = Nothing
            rsFrmSecFollowUp.Close
            Set rsFrmSecFollowUp = Nothing
        End If
        
        
        If xFunction = "ATTENDANCE" Then
            'Add Attendance Codes Security
            Dim rsFrmSecAttend As New ADODB.Recordset
            Dim rsToSecAttend As New ADODB.Recordset
            
            'Retrieve Template's Attendance Security
            SQLQ = "SELECT * FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            rsFrmSecAttend.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            'Delete User's Attendance Security first and then add back based on Template
            SQLQ = "DELETE FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            gdbAdoIhr001.Execute SQLQ
            
            'Open User's Attendance Security record
            SQLQ = "SELECT * FROM HR_SECURE_ATTENDANCE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                
            'Add User's Attendance Security based on the Template
            Do While Not rsFrmSecAttend.EOF
                rsToSecAttend.AddNew
                rsToSecAttend("COMPNO") = "001"
                rsToSecAttend("USERID") = xUserID
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
            Set rsToSecAttend = Nothing
            rsFrmSecAttend.Close
            Set rsFrmSecAttend = Nothing
        End If
        
        
        'Release 8.1
        If xFunction = "DOCUMENTTYPE" Then
            'Add Document Type Codes Security
            Dim rsFrmSecDocType As New ADODB.Recordset
            Dim rsToSecDocType As New ADODB.Recordset
            
            'Retrieve Template's Document Type Security
            SQLQ = "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            rsFrmSecDocType.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            'Delete User's Document Type Security first and then add back based on Template
            SQLQ = "DELETE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            gdbAdoIhr001.Execute SQLQ
            
            'Open User's Document Type Security record
            SQLQ = "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecDocType.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
                
            'Add User's Document Type Security based on the Template
            Do While Not rsFrmSecDocType.EOF
                rsToSecDocType.AddNew
                rsToSecDocType("COMPNO") = "001"
                rsToSecDocType("USERID") = xUserID
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
            Set rsToSecDocType = Nothing
            rsFrmSecDocType.Close
            Set rsFrmSecDocType = Nothing
        End If
        
        'Ticket #30508 - Application Tracking Enhancement
        If xFunction = "REQUISITION" Then
            'Add the Requisition Security
            Dim rsFrmSecRequist As New ADODB.Recordset
            Dim rsToSecRequist As New ADODB.Recordset
            
            'Retrieve Template's Requisition Security
            SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            rsFrmSecRequist.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
            
            'Delete User's Requisition Security first and then add back based on Template
            SQLQ = "DELETE FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            gdbAdoIhr001.Execute SQLQ
            
            'Open User's Requisition Security record
            SQLQ = "SELECT * FROM HRA_SECURE_REQUISITION WHERE USERID='" & Replace(xUserID, "'", "''") & "'"
            rsToSecRequist.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
            
            'Add User's Requisition Security based on the Template
            Do While Not rsFrmSecRequist.EOF
                rsToSecRequist.AddNew
                rsToSecRequist("COMPNO") = "001"
                rsToSecRequist("USERID") = xUserID
                rsToSecRequist("RS_POSTYPE") = rsFrmSecRequist("RS_POSTYPE")
                rsToSecRequist("RS_ORG") = rsFrmSecRequist("RS_ORG")
                rsToSecRequist("RS_GRPCD") = rsFrmSecRequist("RS_GRPCD")
                rsToSecRequist("RS_STATUS") = rsFrmSecRequist("RS_STATUS")
                                
                rsToSecRequist("RS_INCLEMPNBR") = rsFrmSecRequist("RS_INCLEMPNBR")
                rsToSecRequist("RS_EXCLEMPNBR") = rsFrmSecRequist("RS_EXCLEMPNBR")
                
                rsToSecRequist("LDATE") = Date
                rsToSecRequist("LTIME") = Time$
                rsToSecRequist("LUSER") = glbUserID
                rsToSecRequist.Update
                
                rsFrmSecRequist.MoveNext
            Loop
            rsToSecRequist.Close
            Set rsToSecRequist = Nothing
            rsFrmSecRequist.Close
            Set rsFrmSecRequist = Nothing
        End If
        
        
        Screen.MousePointer = DEFAULT
        
'        If xUpdType = "Add" Then
'            MDIMain.panHelp(0).Caption = "Security Add Done"
'            MsgBox "Security Profile added for '" & xUserID & "' successfully.", vbInformation, "Security Added"
'        ElseIf xUpdType = "Update" Then
'            MDIMain.panHelp(0).Caption = "Security Update Done"
'            MsgBox "Security Profile updated for '" & xUserID & "' successfully.", vbInformation, "Security Updated"
'        ElseIf xUpdType = "Reset" Then
'            MDIMain.panHelp(0).Caption = "Security Reset/Update Done"
'            MsgBox "Security Profile has been reset/updated for '" & xUserID & "' successfully based on Security Template '" & xTemplate & "'.", vbInformation, "Security Reset/Update"
'        End If
    End If
End Sub

Public Function Add_Train_FollowUp(xEmpNo, xRenewDate, xCourseCode, xJob) As Long
Dim rsFollowUp As New ADODB.Recordset
Dim SQLQ As String
    
    Add_Train_FollowUp = 0
    
    'Check if Follow Up record already exists - in case of Refreshing the Training Plan. Then update the existing
    'Follow Up record instead of creating a new one.
    SQLQ = "SELECT * FROM HR_FOLLOW_UP"
    SQLQ = SQLQ & " WHERE EF_EMPNBR =" & xEmpNo
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xRenewDate)
    SQLQ = SQLQ & " AND EF_COMMENTS = '" & Replace("Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJob, "'", "''") & "'"
    SQLQ = SQLQ & " AND EF_COMPLETED = 0"
    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsFollowUp.EOF Then
        'Follow Up Record already exists. Do not add a new one. Just update with last update transaction values.
        rsFollowUp("EF_LDATE") = Date
        rsFollowUp("EF_LUSER") = glbUserID
        rsFollowUp("EF_LTIME") = Time$
        rsFollowUp.Update
        
        Add_Train_FollowUp = rsFollowUp("EF_FOLLOWUP_ID")
    Else
        'Add a Follow Up record for this Training course
        'SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE 1 = 2"
        'rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rsFollowUp.AddNew
        rsFollowUp("EF_COMPNO") = "001"
        rsFollowUp("EF_EMPNBR") = xEmpNo
        rsFollowUp("EF_FDATE") = xRenewDate
        rsFollowUp("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsFollowUp("EF_ADMINBY_TABL") = "EDAB"
            rsFollowUp("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsFollowUp("EF_FREAS") = "EDUC"
        rsFollowUp("EF_COMMENTS") = "Course: " & xCourseCode & " - " & GetTABLDesc("ESCD", xCourseCode) & " for Position: " & xJob
        rsFollowUp("EF_LDATE") = Date
        rsFollowUp("EF_LTIME") = Time$
        rsFollowUp("EF_LUSER") = glbUserID
        rsFollowUp.Update
        
        Add_Train_FollowUp = rsFollowUp("EF_FOLLOWUP_ID")
    End If
    rsFollowUp.Close
    Set rsFollowUp = Nothing

End Function

Public Function SMTP_Log(xFunction, xStatus) As Boolean
    
    Dim rsSMTP As New ADODB.Recordset
    Dim SQLQ As String
    
    'Ticket #24629 - Release 8.0
    
    SMTP_Log = False
    
    SQLQ = "SELECT * FROM HR_SMTPLOG WHERE 1 = 2"
    rsSMTP.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsSMTP.AddNew
    rsSMTP("SM_COMPNO") = "001"
    rsSMTP("SM_SOURCE") = Left(MDIMain.ActiveForm.Caption, 50)
    rsSMTP("SM_STATUS") = Left(xStatus, 2000)
    rsSMTP("SM_LDATE") = Date
    rsSMTP("SM_LTIME") = Time$
    rsSMTP("SM_LUSER") = glbUserID
    rsSMTP.Update
    
    rsSMTP.Close
    Set rsSMTP = Nothing
    
    SMTP_Log = True
    
End Function

Public Function HRLog(xEmpNbr, xFDate, xTDate, xAction, xComments, xSource, Optional xTermSEQ) As Boolean
    
    Dim rsHRLog As New ADODB.Recordset
    Dim SQLQ As String
    
    ''Ticket #24805 & Ticket #24485 - Keeping Log for various activities in info:HR.
    'Currently the log is for:
    '           - Multi week Work Schedule
    
    
    HRLog = False
    
    SQLQ = "SELECT * FROM HRLOG WHERE 1 = 2"
    rsHRLog.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsHRLog.AddNew
    rsHRLog("LG_COMPNO") = "001"
    rsHRLog("LG_EMPNBR") = xEmpNbr
    rsHRLog("LG_FDATE") = xFDate
    rsHRLog("LG_TDATE") = xTDate
    rsHRLog("LG_ACTION") = Left(xAction, 20)
    rsHRLog("LG_COMMENTS") = Left(xComments, 4000)
    
    'Make sure the Source Table Code exists in the HRTable
    Call CreateTableMasterCode("ASRC", Left(xSource, 6), Left(xSource, 50))
    rsHRLog("LG_SOURCE") = Left(xSource, 6)
    
    rsHRLog("LG_LDATE") = Date
    rsHRLog("LG_LTIME") = Time$
    rsHRLog("LG_LUSER") = glbUserID
    
    If Not IsMissing(xTermSEQ) Then
        rsHRLog("TERM_SEQ") = xTermSEQ
    End If
    
    rsHRLog.Update
    
    rsHRLog.Close
    Set rsHRLog = Nothing
    
    HRLog = True
    
End Function

Public Sub Grant_FollowUpCode_Security(xUserID, xFCode, xFCodeDesc)
Dim rsSR As New ADODB.Recordset
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(xUserID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    rsSR.Open "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xUserID, "'", "''") & "' AND CODENAME='" & xFCode & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    '????Ticket #24808 -  Retrieve template's security profile
    rsSR.Open "SELECT * FROM HR_SECURE_FOLLOW_UP WHERE USERID='" & Replace(xTemplate, "'", "''") & "' AND CODENAME='" & xFCode & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
If rsSR.EOF Then
    'SQLQ = "INSERT INTO HR_SECURE_FOLLOW_UP(COMPNO,USERID," & Field_SQL("DESCRIPTION") & ",ACCESSABLE,Maintainable,CODENAME, TB_NAME) "
    'SQLQ = SQLQ & " VALUES('001','" & glbSecUSERID & "'," & Chr$(34) & lStr(rsTD("TB_DESC")) & Chr$(34) & ",0,0,'" & rsTD("TB_KEY") & "','ECOM')"
    rsSR.AddNew
    rsSR("COMPNO") = "001"
    rsSR("USERID") = xUserID
    rsSR("DESCRIPTION") = xFCodeDesc
    rsSR("ACCESSABLE") = 1
    rsSR("Maintainable") = 1
    rsSR("CODENAME") = xFCode
    rsSR("TB_NAME") = "FURE"
   rsSR.Update
End If
rsSR.Close
Set rsSR = Nothing

End Sub

Public Sub Grant_DocumentTypeCode_Security(xUserID, xDCode, xDCodeDesc)
Dim rsSR As New ADODB.Recordset
Dim xTemplate As String

'????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
xTemplate = ""
xTemplate = Get_Template(xUserID)

If xTemplate = "" Or xTemplate = "TEMPLATE" Then
    rsSR.Open "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xUserID, "'", "''") & "' AND CODENAME='" & xDCode & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
Else
    '????Ticket #24808 -  Retrieve template's security profile
    rsSR.Open "SELECT * FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xTemplate, "'", "''") & "' AND CODENAME='" & xDCode & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
If rsSR.EOF Then
    'SQLQ = "INSERT INTO HR_SECURE_DOCUMENT_TYPE (COMPNO,USERID," & Field_SQL("DESCRIPTION") & ",ACCESSABLE,Maintainable,CODENAME, TB_NAME) "
    'SQLQ = SQLQ & " VALUES('001','" & glbSecUSERID & "'," & Chr$(34) & lStr(rsTD("TB_DESC")) & Chr$(34) & ",0,0,'" & rsTD("TB_KEY") & "','ECOM')"
    rsSR.AddNew
    rsSR("COMPNO") = "001"
    rsSR("USERID") = xUserID
    rsSR("DESCRIPTION") = xDCodeDesc
    rsSR("ACCESSABLE") = 1
    rsSR("Maintainable") = 1
    rsSR("CODENAME") = xDCode
    rsSR("TB_NAME") = "DOCT"
   rsSR.Update
End If
rsSR.Close
Set rsSR = Nothing

End Sub

Public Sub Update_Age65_LTD_Benefit_EndDate(xEmpNbr, xDOB)
    Dim rsBen As New ADODB.Recordset
    Dim SQLQ As String
    Dim xPer
    Dim xDateAge65
    Dim oEndDate
    Dim nEndDate

    SQLQ = "SELECT * FROM HRBENFT"
    SQLQ = SQLQ & " WHERE BF_EMPNBR = " & xEmpNbr & " AND BF_BCODE = 'LTD'"
    SQLQ = SQLQ & " AND (BF_GROUP <> 'PARTNERS' AND BF_GROUP <> 'ART') "
    rsBen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsBen.EOF Then
        'If IsNumeric(rsBen("BF_WAITPERIOD")) Then
            'xPER = rsBen("BF_DWM")
            'If xPER = "W" Then xPER = "ww"
            
            'Get the date for Age 65
            'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
            'Get the date for Age 65 or Age 67 based on Benefit Group
            If IsNull(rsBen("BF_GROUP")) Or rsBen("BF_GROUP") = "" Then
                xDateAge65 = DateAdd("yyyy", 67, CVDate(xDOB))
            Else
               xDateAge65 = DateAdd("yyyy", 65, CVDate(xDOB))
            End If
            
            'Existing End Date and New End Date
            oEndDate = IIf(IsNull(rsBen("BF_CEASEDATE")), "", rsBen("BF_CEASEDATE"))
            If IsDate(oEndDate) Then
                'nEndDate = MonthLastDate(DateAdd(xPER, 0 - Val(rsBen("BF_WAITPERIOD")), CVDate(xDateAge65)))
                'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
                If IsNull(rsBen("BF_GROUP")) Or rsBen("BF_GROUP") = "" Then
                    nEndDate = CVDate(Format("09/30/" & Year(xDateAge65), "mm/dd/yyyy"))
                Else
                    nEndDate = MonthLastDate(DateAdd("d", 0 - 90, CVDate(xDateAge65)))
                End If
            
                If CVDate(oEndDate) <> CVDate(nEndDate) Then
                    'Compute LTD End Date based on employee's 65th birthday - 90days and get the last date of month
                    'rsBen("BF_CEASEDATE") = MonthLastDate(DateAdd(xPER, 0 - Val(rsBen("BF_WAITPERIOD")), CVDate(xDateAge65)))
                    'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
                    If IsNull(rsBen("BF_GROUP")) Or rsBen("BF_GROUP") = "" Then
                        rsBen("BF_CEASEDATE") = CVDate(Format("09/30/" & Year(xDateAge65), "mm/dd/yyyy"))
                    Else
                        rsBen("BF_CEASEDATE") = MonthLastDate(DateAdd("d", 0 - 90, CVDate(xDateAge65)))
                    End If
                Else
                    rsBen.Close
                    Set rsBen = Nothing
                    Exit Sub
                End If
            Else
                'Compute LTD End Date based on employee's 65th birthday - 90days and get the last date of month
                'rsBen("BF_CEASEDATE") = MonthLastDate(DateAdd(xPER, 0 - Val(rsBen("BF_WAITPERIOD")), CVDate(xDateAge65)))
                'Ticket #27113 - For Partners the Cease Date will be Sept 30th in the year they turn 67
                If IsNull(rsBen("BF_GROUP")) Or rsBen("BF_GROUP") = "" Then
                    rsBen("BF_CEASEDATE") = CVDate(Format("09/30/" & Year(xDateAge65), "mm/dd/yyyy"))
                Else
                    rsBen("BF_CEASEDATE") = MonthLastDate(DateAdd("d", 0 - 90, CVDate(xDateAge65)))
                End If
            End If
            rsBen("BF_LDATE") = Format(Now, "SHORT DATE")
            rsBen("BF_LUSER") = glbUserID
            rsBen("BF_LTIME") = Time$
            rsBen.Update
            
            'Update Audit with End Date
            Call AUDITBENF_EndDate(xEmpNbr, rsBen)
        'End If
    End If
    rsBen.Close
    Set rsBen = Nothing
End Sub

Public Function AUDITBENF_EndDate(xEmpNo, Optional rslBen As ADODB.Recordset)
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDiv As String
Dim strFields As String
Dim ACTX
Dim NBCode, NEDate
Dim xTermSEQ
Dim SQLQ As String

'''On Error GoTo AUDIT_ERR
AUDITBENF_EndDate = False

ACTX = "M"

SQLQ = "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDiv = ""
    Else
        xDiv = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDiv = ""
End If

'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS,AU_CEASEDATE,AU_LDAY "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

xADD = False

NEDate = ""
If Not IsNull(rslBen("BF_EDATE")) Then NEDate = rslBen("BF_EDATE")

'BF_CEASEDATE was changed
MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDiv

rsTA("AU_BCODE") = rslBen("BF_BCODE")
rsTA("AU_EDATE") = rslBen("BF_EDATE")
rsTA("AU_CEASEDATE") = rslBen("BF_CEASEDATE")
rsTA("AU_LDATE") = Date
If IsDate(NEDate) Then 'if benefit effe date is future date, use it as LDATE
    If CVDate(NEDate) > CVDate(Date) Then
        rsTA("AU_LDATE") = CVDate(NEDate)
    End If
End If


SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
Set rsEmp = Nothing

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update
rsTA.Close

MODNOUPD:
AUDITBENF_EndDate = True

Exit Function
AUDIT_ERR:

End Function

Public Function Get_BenefitType_BenefitRateTable(xBenCode)
Dim rsBenRates As New ADODB.Recordset
Dim SQLQ As String

    SQLQ = "SELECT BR_BENTYPE FROM HR_BENEFIT_RATES WHERE BR_BCODE = '" & xBenCode & "'"
    rsBenRates.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsBenRates.EOF Then
        Get_BenefitType_BenefitRateTable = rsBenRates("BR_BENTYPE")
    Else
        Get_BenefitType_BenefitRateTable = ""
    End If
    rsBenRates.Close
    Set rsBenRates = Nothing

End Function

Public Function Get_BenefitRate(xEmpNbr, xBenCode, Optional xBenType As DependentRelationship)
    Dim rsBenRates As New ADODB.Recordset
    
    Dim SQLQ As String
    Dim xAge
    Dim xGender
    Dim xSmoker
        
    Get_BenefitRate = 0
    
    'Get employee/dependent information
    If Not IsMissing(xBenType) Then
        If xBenType = DependentRelationship.Employee Then
            'Get Employee's Information
            xAge = CurrentAge(GetEmpData(xEmpNbr, "ED_DOB"))
            xGender = GetEmpData(xEmpNbr, "ED_SEX")
            xSmoker = GetEmpData(xEmpNbr, "ED_SMOKER")
        Else
            'Get Dependent information
            xAge = CurrentAge(GetDependentData(xEmpNbr, "DP_DOB", IIf(xBenType = Spouse, "Spouse", IIf(xBenType = Children, "Child", "Other"))))
            xGender = GetDependentData(xEmpNbr, "DP_SEX", IIf(xBenType = Spouse, "Spouse", IIf(xBenType = Children, "Child", "Other")))
            xSmoker = GetDependentData(xEmpNbr, "DP_SMOKER", IIf(xBenType = Spouse, "Spouse", IIf(xBenType = Children, "Child", "Other")))
        End If
    Else
        'Get Employee's Information
        xAge = CurrentAge(GetEmpData(xEmpNbr, "ED_DOB"))
        xGender = GetEmpData(xEmpNbr, "ED_SEX")
        xSmoker = GetEmpData(xEmpNbr, "ED_SMOKER")
    End If
    
    'Retrive the Benefit
    SQLQ = "SELECT * FROM HR_BENEFIT_RATES"
    SQLQ = SQLQ & " WHERE BR_AGE_FROM <= " & xAge
    SQLQ = SQLQ & " AND BR_AGE_TO >= " & Int(xAge)      'Ticket #27113 - Allowing to get the rate at any time of the year
    SQLQ = SQLQ & " AND BR_SEX = '" & xGender & "'"
    SQLQ = SQLQ & " AND BR_SMOKER = " & IIf(xSmoker, 1, 0)
    If Not IsMissing(xBenType) Then
        SQLQ = SQLQ & " AND BR_BENTYPE = '" & IIf(xBenType = Spouse, "S", IIf(xBenType = Children, "C", IIf(xBenType = Other, "O", "E"))) & "'"
    Else
        SQLQ = SQLQ & " AND BR_BENTYPE = 'O'"
    End If
    SQLQ = SQLQ & " AND BR_BCODE = '" & xBenCode & "'"
    'SQLQ = SQLQ & " AND BR_RATE = " & medRate.Text
    SQLQ = SQLQ & " ORDER BY BR_AGE_FROM,BR_AGE_TO,BR_SEX,BR_SMOKER,BR_BENTYPE,BR_BCODE"
    rsBenRates.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsBenRates.EOF Then
        Get_BenefitRate = rsBenRates("BR_RATE")
    Else
        Get_BenefitRate = 0
    End If
    rsBenRates.Close
    Set rsBenRates = Nothing
    
End Function

Public Function CurrentAge(xDOB)
    Dim birthdate
    
    'Calculate current Age
    If IsDate(xDOB) Then
        birthdate = CVDate(xDOB)
        
        CurrentAge = DateDiff("m", birthdate, Now)
        If month(birthdate) = month(Now) Then
            If Day(Now) < Day(birthdate) Then
                CurrentAge = CurrentAge - 1
            End If
        End If
        CurrentAge = CDbl(CurrentAge / 12)
    Else
        CurrentAge = 0
    End If

End Function

Public Function GetDependentData(EmpNbr, Field As String, xRelationship, Optional DEFAULT)
    Dim rsDependent As New ADODB.Recordset
    
    'Relations
    'Aunt,Brother,Children,Common Law,Couple,Daughter,Estate,Ex-Spouse,Father,Fiancee,Fiance,Husband,Mother,Other,
    'Parents,Sister,Son,Spouse,Uncle,Wife
    
    rsDependent.Open "SELECT " & Field & " FROM HRDEPEND WHERE DP_EMPNBR=" & EmpNbr & " AND DP_RELATE = '" & xRelationship & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If Not IsMissing(DEFAULT) Then
        GetDependentData = DEFAULT
    End If
    If Not rsDependent.EOF Then
        If Not IsNull(rsDependent(Field)) Then GetDependentData = rsDependent(Field)
    Else
        GetDependentData = ""
    End If
End Function

Public Function getJobFamilyDesc(xCode, xInx) 'Ticket #26233 Franks 11/21/2014 VitalAire Canada Inc.
Dim rsDiv As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = "Unassigned"
    If Not IsNull(xCode) Then
        SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOBFAMILY WHERE JB_CODE = '" & xCode & "' "
        If xInx = 0 Then SQLQ = SQLQ & "AND JB_TYPE = 'JOBFAMILY' "
        If xInx = 1 Then SQLQ = SQLQ & "AND JB_TYPE = 'SUBFAMILY' "
        If xInx = 2 Then SQLQ = SQLQ & "AND JB_TYPE = 'GROUPJOBS' "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            xRetVal = rsDiv("JB_DESCR")
        End If
        rsDiv.Close
    End If
    getJobFamilyDesc = xRetVal
End Function

Public Sub SetCurrentSalary_OFF(xEmpNbr, xPosCode, xPosStartDate)
    Dim rsSalHis As New ADODB.Recordset
    Dim SQLQ As String
    
    'Retrieve and Update Current Salary record matching the Position and Start Date whose Current Flag just got turned OFF.
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
    SQLQ = SQLQ & " WHERE SH_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND SH_JOB = '" & xPosCode & "'"
    SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(xPosStartDate)
    SQLQ = SQLQ & " AND SH_CURRENT <> 0"
    rsSalHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsSalHis.EOF
        rsSalHis("SH_CURRENT") = False
        rsSalHis("SH_PRIMARY") = False
        rsSalHis("SH_LDATE") = Date
        rsSalHis("SH_LTIME") = Time$
        rsSalHis("SH_LUSER") = glbUserID
        rsSalHis.Update
        
        rsSalHis.MoveNext
    Loop
    rsSalHis.Close
    Set rsSalHis = Nothing
End Sub

Public Function UpdatePrimaryPositionSalary(xEmpNbr)
Dim rsEmpJob As New ADODB.Recordset
Dim rsSalHis As New ADODB.Recordset
Dim SQLQ As String

    'Turn OFF Primary Position on all Salary records before updating with correct one
    SQLQ = "UPDATE HR_SALARY_HISTORY SET SH_PRIMARY = 0"
    SQLQ = SQLQ & " WHERE SH_EMPNBR=" & xEmpNbr
    gdbAdoIhr001.Execute SQLQ

    'Retrieve Primary Position and Update Salary
    SQLQ = "SELECT JH_EMPNBR, JH_ID, JH_PRIMARY, JH_JOB, JH_SDATE FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNbr & " "
    SQLQ = SQLQ & " AND JH_PRIMARY <> 0 "
    SQLQ = SQLQ & " AND JH_CURRENT <> 0"
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpJob.EOF Then
        SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
        SQLQ = SQLQ & " WHERE SH_EMPNBR = " & xEmpNbr
        SQLQ = SQLQ & " AND SH_JOB = '" & rsEmpJob("JH_JOB") & "'"
        SQLQ = SQLQ & " AND SH_SDATE = " & Date_SQL(rsEmpJob("JH_SDATE"))
        SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_CURRENT " & IIf(glbSQL Or glbOracle, "DESC", "")
        rsSalHis.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsSalHis.EOF Then
            rsSalHis.MoveFirst
            
            rsSalHis("SH_CURRENT") = True
            rsSalHis("SH_PRIMARY") = True
            rsSalHis("SH_LDATE") = Date
            rsSalHis("SH_LTIME") = Time$
            rsSalHis("SH_LUSER") = glbUserID
            rsSalHis.Update
            
            rsSalHis.MoveNext
        End If
        rsSalHis.Close
        Set rsSalHis = Nothing
    End If
    rsEmpJob.Close
    Set rsEmpJob = Nothing
End Function

Public Function Update_Attendance_SalaryInfo(rsNew As ADODB.Recordset)
    Dim rsAttend As New ADODB.Recordset
    Dim SQLQ As String
    
    'Update employee's Attendance records with the new salary based on the Salary Effective Date and matching
    'Position code and in the Attendance record.
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsNew("SH_EMPNBR")
    SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(rsNew("SH_EDATE"))
    SQLQ = SQLQ & " AND AD_JOB = '" & rsNew("SH_JOB") & "'"
    SQLQ = SQLQ & " AND AD_SALCD ='" & rsNew("SH_SALCD") & "'"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttend.EOF
        rsAttend("AD_SALARY") = rsNew("SH_SALARY")
        
        rsAttend("AD_LDATE") = Date
        rsAttend("AD_LUSER") = glbUserID
        rsAttend("AD_LTIME") = Time$
        rsAttend.Update
        
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing
End Function

Public Function Update_Attendance_SalaryInfo1(xEmpNbr, xJob, xEDate, xSalary, xPer)
    Dim rsAttend As New ADODB.Recordset
    Dim SQLQ As String
    
    'Update employee's Attendance records with the new salary based on the Salary Effective Date and matching
    'Position code and in the Attendance record.
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNbr
    SQLQ = SQLQ & " AND AD_DOA >= " & Date_SQL(xEDate)
    SQLQ = SQLQ & " AND AD_JOB = '" & xJob & "'"
    SQLQ = SQLQ & " AND AD_SALCD ='" & xPer & "'"
    rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsAttend.EOF
        rsAttend("AD_SALARY") = xSalary
        
        rsAttend("AD_LDATE") = Date
        rsAttend("AD_LUSER") = glbUserID
        rsAttend("AD_LTIME") = Time$
        rsAttend.Update
        
        rsAttend.MoveNext
    Loop
    rsAttend.Close
    Set rsAttend = Nothing
End Function

Public Function Update_Employee_With_DailyAccrual()
    Dim SQLQ As String
    Dim rsDailyEnt As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim rsDailyAcc As New ADODB.Recordset
    Dim lngRecs As Long
    Dim dblDHours#, intWhereFit&, x%, dblNewEntitle#
    Dim dblFTEHours#, dblWHours#
    Dim dblServiceYears#
    Dim if_Entitle As Boolean
    Dim xComments As String
    Dim recNo As Long
    Dim xAsOf
    Dim pct
    Dim varStartDate
    Dim dblEntitle#, dblNewDailyEnt, dblEntitleUpd#
    Dim xTotEmpHours
    Dim xUpdated As Boolean
    
    xUpdated = False
    
    'Check first if rules exists as the client may be just setting up the Daily Entitlement rules so no need to run this at this moment.
    If Check_Daily_Entitlement_Rule_Exists And Check_Daily_Accrual_Exists("1=1") Then
    
        'For each rule - get the list of employees and then retrieve their daily entitlement to update.
        SQLQ = "SELECT DISTINCT VD_ORG,VD_EMP,VD_EMPEXCL,VD_PT,VD_EDATE,VD_FRDATE,VD_TODATE,VD_MANUAL FROM HRVACENTDAILY ORDER BY VD_FRDATE,VD_TODATE"
        rsDailyEnt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsDailyEnt.EOF Then
        
            MsgBox "Vacation accruals are being updated. Your login will continue once this process is completed.", vbInformation, "Daily Vacation Accrual Update"
            
            rsDailyEnt.MoveFirst
            
            Do While Not rsDailyEnt.EOF
                'Get the Employees who had entitlement daily accrual files
                SQLQ = "SELECT ED_EMPNBR, ED_ORG, ED_EMP, ED_PT, ED_VAC,ED_PVAC,ED_VACT, ED_ANNVAC, ED_EFDATE, ED_ETDATE FROM HREMP WHERE "
                SQLQ = SQLQ & " ED_EFDATE <> '' AND ED_EFDATE IS NOT NULL AND ED_ETDATE <> '' AND ED_ETDATE IS NOT NULL"
                
                'Build Selection criteria based on the rule
                If Len(rsDailyEnt("VD_ORG")) > 0 Then
                    SQLQ = SQLQ & " AND ED_ORG = '" & rsDailyEnt("VD_ORG") & "'"
                End If
                If Len(rsDailyEnt("VD_EMP")) > 0 Then
                    SQLQ = SQLQ & " AND ED_EMP = '" & rsDailyEnt("VD_EMP") & "'"
                End If
                If Len(rsDailyEnt("VD_PT")) > 0 Then
                    SQLQ = SQLQ & " AND ED_PT = '" & rsDailyEnt("VD_PT") & "' "
                End If
                If Len(rsDailyEnt("VD_EMPEXCL")) > 0 Then
                    SQLQ = SQLQ & " AND (ED_EMP NOT IN ('" & Replace(rsDailyEnt("VD_EMPEXCL"), ",", "','") & "'))"
                End If
                If IsDate(rsDailyEnt("VD_EDATE")) Then
                    SQLQ = SQLQ & " AND ED_EFDATE >= " & Date_SQL(rsDailyEnt("VD_EDATE"))
                End If
                If IsDate(rsDailyEnt("VD_FRDATE")) Then
                    SQLQ = SQLQ & " AND ED_EFDATE = " & Date_SQL(rsDailyEnt("VD_FRDATE"))
                End If
                If IsDate(rsDailyEnt("VD_TODATE")) Then
                    SQLQ = SQLQ & " AND ED_ETDATE = " & Date_SQL(rsDailyEnt("VD_TODATE"))
                End If
                
                rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If rsHREmp.BOF And rsHREmp.EOF Then
                    'MsgBox "Employees for this selection do not exist!"
                    'Start_Processing = False
                    GoTo nextEntRule
                Else
                    'MsgBox "Vacation accruals are being updated. Your login will continue once this process is completed.", vbInformation, "Daily Vacation Accrual Update"
                    
                    'Update employee's Vacation with accruals upto current day
                    lngRecs = rsHREmp.RecordCount
                    
                    rsHREmp.MoveFirst
                    
                    MDIMain.panHelp(0).Caption = ""
                    MDIMain.panHelp(0).FloodType = 1
                    MDIMain.panHelp(0).FloodPercent = 5
                    MDIMain.panHelp(1).Caption = "Daily Vacation Entitlement Update"
                    
                    'gdbAdoIhr001.BeginTrans
                    recNo = 0
                    
                    'For each employee create the daily accrual
                    While Not rsHREmp.EOF
                        recNo = recNo + 1
                        pct = Int(100 * (recNo / lngRecs))
                        MDIMain.panHelp(0).FloodPercent = pct
                    
                        'Check if the Daily Accrual details exists for this employee. If not skip to next employee
                        SQLQ = " DA_EMPNBR = " & rsHREmp("ED_EMPNBR")
                        SQLQ = SQLQ & " AND DA_FRDATE = " & Date_SQL(rsHREmp("ED_EFDATE"))
                        SQLQ = SQLQ & " AND DA_TODATE = " & Date_SQL(rsHREmp("ED_ETDATE"))
                        If Check_Daily_Accrual_Exists(SQLQ) Then
                        
                            If IsNull(rsHREmp("ED_VAC")) Then
                                dblEntitle# = 0
                            Else
                                dblEntitle# = rsHREmp("ED_VAC")
                            End If
                                                
                            'Start with 0 entitlement if not 0
                            'If Not IsNull(rsHREmp("ED_VAC")) And rsHREmp("ED_VAC") <> "" And Val(rsHREmp("ED_VAC")) <> 0 Then
                            '    'Update Accrual Table as well
                            '    xComments = "Current Vac. Ent. Chg from " & rsHREmp("ED_VAC") & " to 0"
                            '    Call Append_Accrual(rsHREmp("ED_EMPNBR"), "VAC", CVDate(Format(dlpDateRange(0).Text, "mm/dd/yyyy")), -Val(rsHREmp("ED_VAC") & ""), "Z", xComments)
                            '
                            '    rsHREmp("ED_VAC") = 0
                            '    rsHREmp("ED_ANNVAC") = 0
                            '    rsHREmp.Update
                            'End If
                            
                            'Start from Vacation Entitlement Start Date
                            xAsOf = rsHREmp("ED_EFDATE")
                            
                            'For each day from Start of the Vacation Entitlement period to Current Day, get employee's daily accrual and update employee's Vacation entitlement
                            Do While CVDate(xAsOf) <= CVDate(Date)
                                'Retrieve day's accrual from Daily Accrual file
                                dblEntitleUpd# = Round(Get_DailyAccrual(rsHREmp("ED_EMPNBR"), rsDailyEnt("VD_ORG"), rsDailyEnt("VD_EMP"), rsDailyEnt("VD_PT"), "", rsHREmp("ED_EFDATE"), rsHREmp("ED_ETDATE"), xAsOf, False), 4)
                            
                                'Only update if not 0
                                If dblEntitleUpd# <> 0 Then
                                    'Update ED_VAC in HREMP table with day's accrual (from begining of the entitlement period to today's date)
                                    rsHREmp("ED_VAC") = dblEntitle# + dblEntitleUpd
                                    rsHREmp.Update
                                                                
                                    'Update the Daily Accrual table with Process Date as this will append to employee's Vacation Entitlement
                                    Call DailyAccrual_Processed(rsHREmp("ED_EMPNBR"), rsDailyEnt("VD_ORG"), rsDailyEnt("VD_EMP"), rsDailyEnt("VD_PT"), "", rsHREmp("ED_EFDATE"), rsHREmp("ED_ETDATE"), xAsOf, CVDate(Format(Date, "mm/dd/yyyy")))
                                    
                                    'Append in Accrual table as well
                                    xComments = "Current Vac. Ent. Chg from " & dblEntitle# & " to " & rsHREmp("ED_VAC") & ". OS: " & (IIf(IsNull(rsHREmp("ED_PVAC")), 0, rsHREmp("ED_PVAC")) + IIf(IsNull(rsHREmp("ED_VAC")), 0, rsHREmp("ED_VAC"))) - IIf(IsNull(rsHREmp("ED_VACT")), 0, rsHREmp("ED_VACT"))
                                    Call Append_Accrual(rsHREmp("ED_EMPNBR"), "VAC", CVDate(Format(xAsOf, "mm/dd/yyyy")), dblEntitleUpd#, "U", xComments)
                                
                                    'Update local variable with the new ED_VAC so it can be used for next day for accumulation (above)
                                    dblEntitle# = rsHREmp("ED_VAC")
                                End If
lblNextDay:
                                'Move to Next day of the Vacation Entitlement Period
                                xAsOf = DateAdd("d", 1, CVDate(xAsOf))
                            Loop
                        End If
lblNextRec:
                        rsHREmp.MoveNext
                        DoEvents
                        
                        'For updating HRPARCO with last Daily Vacation Entitlemeent Update date
                        xUpdated = True
                    Wend
                End If
nextEntRule:
                rsHREmp.Close
                Set rsHREmp = Nothing

                rsDailyEnt.MoveNext
            Loop
            
            'For updating HRPARCO with last Daily Vacation Entitlemeent Update date
            xUpdated = True
            
        End If
        rsDailyEnt.Close
        Set rsDailyEnt = Nothing
    End If
    
    'Start_Processing = True
    
    'Update HRPARCO with the last update date so this routine is not run again today when the next time anyone logs in
    If xUpdated Then
        SQLQ = "UPDATE HRPARCO SET PC_LST_DAILYVAC_UPD_DATE = " & Date_SQL(Date)
        gdbAdoIhr001.Execute SQLQ
    End If
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    
End Function

Public Sub DailyAccrual_Processed(xEmpNo, xORG, xEMP, xPT, xEmpExclude, xFromDate, xToDate, xAccDate, xProcessDate)
    Dim rsDailyAcc As New ADODB.Recordset
    Dim SQLQ
    
    'Update Day's Accrual as Processed
    SQLQ = "SELECT * FROM HR_DAILYVACACCR "
    SQLQ = SQLQ & " WHERE DA_EMPNBR = " & xEmpNo
    SQLQ = SQLQ & " AND DA_FRDATE = " & Date_SQL(xFromDate)
    SQLQ = SQLQ & " AND DA_TODATE = " & Date_SQL(xToDate)
    SQLQ = SQLQ & " AND DA_ACCRDATE = " & Date_SQL(xAccDate)
    If Len(xORG) = 0 Then
        SQLQ = SQLQ & " AND (DA_ORG IS NULL OR DA_ORG='') "
    Else
        SQLQ = SQLQ & " AND DA_ORG = '" & xORG & "'"
    End If
    If Len(xEMP) = 0 Then
        SQLQ = SQLQ & " AND (DA_EMP IS NULL OR DA_EMP='')"
    Else
        SQLQ = SQLQ & " AND DA_EMP = '" & xEMP & "'"
    End If
    If Len(xPT) = 0 Then
        SQLQ = SQLQ & " AND (DA_PT IS NULL OR DA_PT='')"
    Else
        SQLQ = SQLQ & " AND DA_PT = '" & xPT & "' "
    End If
    If Len(xEmpExclude) = 0 Then
        SQLQ = SQLQ & " AND (DA_EMPEXCL IS NULL OR DA_EMPEXCL='')"
    Else
        SQLQ = SQLQ & " AND DA_EMPEXCL = '" & xEmpExclude & "'"
    End If
    rsDailyAcc.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsDailyAcc("DA_PROCESSDATE") = xProcessDate
    rsDailyAcc("DA_LUSER") = glbUserID
    rsDailyAcc("DA_LDATE") = Date
    rsDailyAcc("DA_LTIME") = Time$
    rsDailyAcc.Update
    rsDailyAcc.Close
    Set rsDailyAcc = Nothing
    
End Sub

Public Function Recompute_DailyAccrualFile(xEmpNbr, xEffDate)
    Dim SQLQ As String
    Dim SQLQ1 As String
    Dim SQLQV As String
    Dim xFromDate, xToDate
    
    'Updated the employee with the new Entitlement period incase it had not been updated yet.
    'This function also computes the TAKEN based on the new entitlement period.
    Call EntReCalcPeriod_Daily("ED_EMPNBR = " & xEmpNbr, "VAC")
    
    'Get Entitlement Period of the employee
    xFromDate = GetEmpData(xEmpNbr, "ED_EFDATE")
    xToDate = GetEmpData(xEmpNbr, "ED_ETDATE")
    
    'If the Effective Date is within the entitlement period then only Daily Accrual update will be needed.
    If IsDate(xFromDate) And IsDate(xToDate) Then
        If CVDate(xEffDate) >= CVDate(xFromDate) And CVDate(xEffDate) <= CVDate(xToDate) Then
            'Get the rule the employee belongs to so we can check if there is Accrual Details exists for this employee
            SQLQV = Get_Employees_DailyEntitlement_Rule(xEmpNbr, xFromDate, xToDate)
            If Len(SQLQV) > 0 Then
                SQLQV = SQLQV & " AND DA_EMPNBR = " & xEmpNbr
            End If
            
            'Accrual details selection
            SQLQ = " DA_EMPNBR = " & xEmpNbr
            SQLQ = SQLQ & " AND DA_FRDATE = " & Date_SQL(xFromDate)
            SQLQ = SQLQ & " AND DA_TODATE = " & Date_SQL(xToDate)
            
            'Clear existing Daily Accruals from the Effective Date onwards
            SQLQ1 = " AND DA_ACCRDATE >= " & Date_SQL(xEffDate)
            
            If Len(SQLQV) > 0 Then
                If Check_Daily_Accrual_Exists(SQLQV) Then
                    'Daily Accrual exists for the employee
                    
                    'This function will clear the ED_VAC as well and update HR_ACCRUAL table. Later on in these steps ED_VAC will be updated again.
                    'Only clearing based on the Entitlement Period and Effective Date. The employee can be part of other rule before and had accrual details for that,
                    'so we need to clear those ones from Effective Date forward.
                    If Clear_Employees_Daily_Accruals(SQLQ & SQLQ1, xEffDate) Then
                        'Daily Accrual cleared successful
                        'Create new Daily Accruals for the period deleted
                        Call Create_Daily_Accrual_File("(ED_EMPNBR = " & xEmpNbr & ")", xEffDate, xFromDate, xToDate)
                        
                        'Update employee's ED_VAC with the new entitlement up to current date. This may include the accrual from previous rule. Therefore,
                        'only going based on the entitlement period
                        Call EntRecalVacDaily(SQLQ, xEffDate)
                    End If
                Else
                    'Clear the Daily Accrual details of the employee if he/she belonged to any other rules for this period.
                    'This function will clear the ED_VAC as well and update HR_ACCRUAL table. Later on in these steps ED_VAC will be updated again.
                    'Only clearing based on the Entitlement Period and Effective Date. The employee can be part of other rule before and had accrual details for that,
                    'so we need to clear those ones from Effective Date forward.
                    If Clear_Employees_Daily_Accruals(SQLQ & SQLQ1, xEffDate) Then
                    
                        'Daily Accrual details do not exists for the employee - Create one and update the employee.
                        Call Create_Daily_Accrual_File("(ED_EMPNBR = " & xEmpNbr & ")", xEffDate, xFromDate, xToDate)
                        
                        'Update employee's ED_VAC & ED_ANNACC with the new entitlement up to current date. This may include the accrual from previous rule. Therefore,
                        'only going based on the entitlement period
                        Call EntRecalVacDaily(SQLQ, xEffDate)
                    End If
                End If
            Else
                'Employee do not belong to any rules
                                
                'Clear any Daily Accruals from Effective Date forward if it exists from the older rule
                SQLQ = "DELETE FROM HR_DAILYVACACCR WHERE " & SQLQ & SQLQ1
                gdbAdoIhr001.BeginTrans
                gdbAdoIhr001.Execute SQLQ
                gdbAdoIhr001.CommitTrans
                
                'Update employee's ED_VAC with the new entitlement up to < Effective date. This may include the accrual from previous rule. Therefore,
                'only going based on the entitlement period
                'Accrual details selection
                SQLQ = " DA_EMPNBR = " & xEmpNbr
                SQLQ = SQLQ & " AND DA_FRDATE = " & Date_SQL(xFromDate)
                SQLQ = SQLQ & " AND DA_TODATE = " & Date_SQL(xToDate)
                Call EntRecalVacDaily(SQLQ, xEffDate)
            End If
        End If
    End If
    
End Function

Public Function Create_Daily_Accrual_File(xSQLQ, xEffDate, xFromDate, xToDate)
    Dim SQLQ As String
    Dim SQLQV As String
    Dim rsDailyVacEnt As New ADODB.Recordset
    Dim rsDailyVacEntlmt As New ADODB.Recordset
    Dim rsHREmp As New ADODB.Recordset
    Dim rsDailyAcc As New ADODB.Recordset
    Dim rsJOB As New ADODB.Recordset
    Dim lngRecs As Long
    Dim dblDHours#, intWhereFit&, x%, dblNewEntitle#
    Dim dblFTEHours#, dblWHours#
    Dim dblServiceYears#
    Dim if_Entitle As Boolean
    Dim xComments As String
    Dim recNo As Long
    Dim xAsOf
    Dim pct
    Dim varStartDate
    Dim dblEntitle#, dblNewDailyEnt, dblEntitleUpd#, xDayB4AccToDate, xLstAccToDate
    Dim xTotEmpHours
    Dim xORG, xEMP, xEmpMode, xEmpExcl
    Dim fglbWDate$
    Dim lstAnnEnt
    Dim flgLvlChanged, flgMidYearStart
        
    'Get the current selected rule
    'Call getWSQLQ("")
       
    Select Case glbCompWDate$ ' sets field reference for basic 'which date'
        Case "O": fglbWDate$ = "ED_DOH"
        Case "S": fglbWDate$ = "ED_SENDTE"
        Case "U": fglbWDate$ = "ED_UNION"
        Case "L": fglbWDate$ = "ED_LTHIRE"
        Case "D": fglbWDate$ = "ED_USRDAT1"
    End Select
       
    'Find the rule the employee belongs to and then create/re-create the Accrual Details for that rule for that employee
    SQLQ = "SELECT DISTINCT VD_ORG,VD_EMP,VD_EMPEXCL,VD_PT,VD_FRDATE,VD_TODATE,VD_MANUAL,VD_EDATE FROM HRVACENTDAILY ORDER BY VD_FRDATE,VD_TODATE"
    rsDailyVacEnt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDailyVacEnt.EOF Then
        rsDailyVacEnt.MoveFirst
        Do While Not rsDailyVacEnt.EOF
        
            'Selection criteria
            xORG = rsDailyVacEnt("VD_ORG") & ""
            xEMP = rsDailyVacEnt("VD_EMP") & ""
            xEmpMode = rsDailyVacEnt("VD_PT") & ""
            xEmpExcl = rsDailyVacEnt("VD_EMPEXCL") & ""
                
            'SQLQV = " AND (ED_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR ED_ORG IS NULL ", "") & ")"
            'SQLQV = SQLQV & " AND (ED_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR ED_EMP IS NULL ", "") & ")"
            'SQLQV = SQLQV & " AND (ED_PT = '" & xEmpMode & "' " & IIf(Len(xEmpMode) = 0, " OR ED_PT IS NULL ", "") & ")"
            'If Len(xEmpExcl) > 0 Then SQLQV = SQLQV & " AND (ED_EMP NOT IN ('" & Replace(xEmpExcl, ",", "','") & "'))"
            
            SQLQV = ""
            If Len(rsDailyVacEnt("VD_ORG")) > 0 Then
                SQLQV = SQLQV & " AND ED_ORG = '" & rsDailyVacEnt("VD_ORG") & "'"
            End If
            If Len(rsDailyVacEnt("VD_EMP")) > 0 Then
                SQLQV = SQLQV & " AND ED_EMP = '" & rsDailyVacEnt("VD_EMP") & "'"
            End If
            If Len(rsDailyVacEnt("VD_PT")) > 0 Then
                SQLQV = SQLQV & " AND ED_PT = '" & rsDailyVacEnt("VD_PT") & "' "
            End If
            If Len(rsDailyVacEnt("VD_EMPEXCL")) > 0 Then
                SQLQV = SQLQV & " AND (ED_EMP NOT IN ('" & Replace(rsDailyVacEnt("VD_EMPEXCL"), ",", "','") & "'))"
            End If
            If IsDate(rsDailyVacEnt("VD_EDATE")) Then
                SQLQV = SQLQV & " AND ED_EFDATE >= " & Date_SQL(rsDailyVacEnt("VD_EDATE"))
            End If
            If IsDate(rsDailyVacEnt("VD_FRDATE")) Then
                SQLQV = SQLQV & " AND ED_EFDATE = " & Date_SQL(rsDailyVacEnt("VD_FRDATE"))
            End If
            If IsDate(rsDailyVacEnt("VD_TODATE")) Then
                SQLQV = SQLQV & " AND ED_ETDATE = " & Date_SQL(rsDailyVacEnt("VD_TODATE"))
            End If
        
    
            'Get the Employees for whom to create the daily accrual files
            SQLQ = "SELECT ED_EMPNBR, ED_VAC,ED_PVAC,ED_VACT, ED_EFDATE, ED_ETDATE, ED_DOH, ED_SENDTE, ED_UNION, ED_LTHIRE, ED_USRDAT1 FROM HREMP WHERE " & xSQLQ & SQLQV & " "
            rsHREmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsHREmp.BOF And rsHREmp.EOF Then
                'Employee do not belong to this rule go to Next Rule
                GoTo Next_Rule
            Else
                'Create Daily Accrual File
                lngRecs = rsHREmp.RecordCount
                
                rsHREmp.MoveFirst
                
                MDIMain.panHelp(0).FloodType = 1
                MDIMain.panHelp(0).FloodPercent = 5
                                
                'Retrieve the Complete Daily Entitlement Rule of the employee
                SQLQ = "SELECT * FROM HRVACENTDAILY WHERE "
                SQLQ = SQLQ & " (VD_ORG = '" & xORG & "' " & IIf(Len(xORG) = 0, " OR VD_ORG IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VD_EMP = '" & xEMP & "' " & IIf(Len(xEMP) = 0, " OR VD_EMP IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VD_PT = '" & xEmpMode & "' " & IIf(Len(xEmpMode) = 0, " OR VD_PT IS NULL ", "") & ")"
                SQLQ = SQLQ & " AND (VD_EMPEXCL = '" & xEmpExcl & "' " & IIf(Len(xEmpExcl) = 0, " OR VD_EMPEXCL IS NULL ", "") & ")"
                SQLQ = SQLQ & " ORDER BY VD_FRDATE,VD_TODATE, VD_ORDER"
                rsDailyVacEntlmt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsDailyVacEntlmt.EOF Then
                    'gdbAdoIhr001.BeginTrans
                    recNo = 0
                    
                    'For each employee create the daily accrual
                    While Not rsHREmp.EOF
                        recNo = recNo + 1
                        pct = Int(100 * (recNo / lngRecs))
                        MDIMain.panHelp(0).FloodPercent = pct
                    
                        If IsNull(rsHREmp("ED_VAC")) Then
                            dblEntitle# = 0
                        Else
                            dblEntitle# = rsHREmp("ED_VAC")
                        End If
                    
                        'Employee's Date used to compute service range against
                        If IsNull(rsHREmp(fglbWDate$)) Then GoTo lblNextRec     'Employee's Entitlement Mass Update Based On Date missing - skip the employee
                        varStartDate = rsHREmp(fglbWDate$)
                        
                        'Get Hours/Day, FTE and Hours/Week
                        If rsJOB.State <> 0 Then rsJOB.Close
                        rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & rsHREmp("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
                        dblDHours# = 0
                        dblFTEHours# = 0
                        dblWHours# = 0
                        If Not rsJOB.EOF Then
                            If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
                            If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
                            dblWHours# = GetJHData(rsHREmp("ED_EMPNBR"), "JH_WHRS", 0)
                        End If
                                    
                        If glbLinamar Then dblDHours# = 8
                                    
                        'Initialise
                        lstAnnEnt = 0
                        flgLvlChanged = False
                        flgMidYearStart = False
                                    
                        'Start from Vacation Entitlement Start Date
                        'xAsOf = xEffDate
                        'Start from Vacation Entitlement Start Date unless Vacation/Sick Mass Update Based Upon Date is greater than Entitlement Start Date
                        If IsDate(rsHREmp(fglbWDate$)) Then
                            If CVDate(rsHREmp(fglbWDate$)) > CVDate(xEffDate) Then
                                xAsOf = rsHREmp(fglbWDate$)
                                flgMidYearStart = True
                            Else
                                xAsOf = xEffDate
                                flgMidYearStart = False
                            End If
                        Else
                            xAsOf = xEffDate
                            flgMidYearStart = False
                        End If
                        
                        'For each day from Start of the Vacation Entitlement period to End Date, compute the daily accrual
                        Do While CVDate(xAsOf) <= CVDate(xToDate)
                        
                            'Compute # of service months
                            dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
                            
                            'Initialize
                            intWhereFit& = -1
                            if_Entitle = False
                            
                            'Which range of the service month the employee falls in and if the entitlement exists for that range
                            rsDailyVacEntlmt.MoveFirst
                            Do While Not rsDailyVacEntlmt.EOF
                                If rsDailyVacEntlmt("VD_EMONTH") > 0 Then
                                    If dblServiceYears# >= CDbl(rsDailyVacEntlmt("VD_BMONTH")) And dblServiceYears# <= CDbl(rsDailyVacEntlmt("VD_EMONTH")) Then
                                        intWhereFit& = 1
                                        If Len(rsDailyVacEntlmt("VD_ENTITLE")) > 0 Then if_Entitle = True
                                        Exit Do
                                    End If
                                End If
                                rsDailyVacEntlmt.MoveNext
                            Loop
                            
                            If intWhereFit& = -1 Then
                                'Skip to next day if not in any of the ranges but first update the Skipped table for audit
                                'Employee #, Status, Union, Category, Excluded Status, Hours/Day, FTE, Date Skipped, Accrual Missed, Reason
                                Call Log_Skipped_Transaction(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), rsDailyVacEntlmt("VD_FRDATE"), rsDailyVacEntlmt("VD_TODATE"), dblDHours#, dblFTEHours#, xAsOf, "", "No Annual Accrual found for " & dblServiceYears# & " Service month")
                                                    
                                'Add the daily accrual to the Daily Accrual details table, as Skipped Day
                                Call Append_Daily_Accrul_File(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), xEffDate, rsDailyVacEntlmt("VD_FRDATE"), rsDailyVacEntlmt("VD_TODATE"), 0, xAsOf, 0, "", Accrued_ToDate(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), rsDailyVacEntlmt("VD_FRDATE"), rsDailyVacEntlmt("VD_TODATE"), xAsOf) + 0, True)
                                                    
                                GoTo lblNextDay
                            End If
                            
                            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
                            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012, they need the total of hours for multiple current positions
                                xTotEmpHours = 0
                                Do While Not rsJOB.EOF
                                    If rsDailyVacEntlmt("VD_TYPE") = "D" Then  ' Entitlements entered in days
                                        If IsNumeric(rsJOB("JH_DHRS")) Then xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS")
                                    End If
                                    If rsDailyVacEntlmt("VD_TYPE") = "F" Then  ' FTE
                                        If IsNumeric(rsJOB("JH_DHRS")) And IsNumeric(rsJOB("JH_FTENUM")) Then
                                            xTotEmpHours = xTotEmpHours + rsJOB("JH_DHRS") * rsJOB("JH_FTENUM")
                                        End If
                                    End If
                                    rsJOB.MoveNext
                                Loop
                            End If
                        
                            'Compute Daily Accrual
                            If if_Entitle Then
                                'Annual Accrual entitled as per the Service range
                                dblNewEntitle# = rsDailyVacEntlmt("VD_ENTITLE")
                            
                                'If Annual Accual is based on Day or FTE and employee is missing these then skip the employee and update the Skipped table
                                'Employee #, Status, Union, Category, Excluded Status, Hours/Day, FTE, Date Skipped, Accrual Missed, Reason
                                If rsDailyVacEntlmt("VD_TYPE") = "D" Then
                                    If dblDHours# = 0 Then
                                        Call Log_Skipped_Transaction(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), rsDailyVacEntlmt("VD_FRDATE"), rsDailyVacEntlmt("VD_TODATE"), dblDHours#, dblFTEHours#, xAsOf, dblNewEntitle#, "Employee's Hours per Day is missing; cannot compute Daily Accrual")
                                        
                                        GoTo lblNextRec
                                    End If
                                End If
                                If rsDailyVacEntlmt("VD_TYPE") = "F" Then
                                    If dblFTEHours# = 0 Then
                                        Call Log_Skipped_Transaction(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), rsDailyVacEntlmt("VD_FRDATE"), rsDailyVacEntlmt("VD_TODATE"), dblDHours#, dblFTEHours#, xAsOf, dblNewEntitle#, "Employee's or FTE is missing; cannot compute Daily Accrual")
                                        
                                        GoTo lblNextRec
                                    End If
                                End If
                                
                                'Annual Accruals in Days to Hours
                                If rsDailyVacEntlmt("VD_TYPE") = "D" Then
                                    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
                                    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
                                        dblNewEntitle# = dblNewEntitle# * xTotEmpHours
                                    Else
                                        dblNewEntitle# = dblNewEntitle# * dblDHours#
                                    End If
                                End If
                                
                                'Annual Accruals by FTE to Hours
                                If rsDailyVacEntlmt("VD_TYPE") = "F" Then
                                    'Ticket #22766 - KidsLink - sum up the FTE for multi positions
                                    If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then  'Kerrys Place Ticket #21843 Franks 04/12/2012
                                        dblNewEntitle# = dblNewEntitle# * xTotEmpHours
                                    Else
                                        dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
                                    End If
                                End If
                                
                                'Annual Accruals in Hours
                                If rsDailyVacEntlmt("VD_TYPE") = "H" Then
                                    'No conversion required as the accruals are stored in hours anyways
                                    dblNewEntitle# = dblNewEntitle#
                                End If
                                
                                'Routine to check if employee moved from one level to another
                                If lstAnnEnt = 0 Then
                                    lstAnnEnt = dblNewEntitle#
                                    flgLvlChanged = False
                                ElseIf lstAnnEnt <> dblNewEntitle# Then
                                    lstAnnEnt = dblNewEntitle#
                                    flgLvlChanged = True
                                End If
                                
                                'Convert the Annual Accrual to Daily Accrual
                                '(1 day * annually earned accrual hours per year) / 365 days (rounded to 4 decimals)
                                'Leap year (year is evenly divided by 4 with no remainder), the number of days for the year will be 366
                                If GetLeapYear(Year(Date)) Then
                                    dblNewDailyEnt = Round((1 * dblNewEntitle) / 366, 4)
                                Else
                                    dblNewDailyEnt = Round((1 * dblNewEntitle) / 365, 4)
                                End If
                                                    
                                'Commenting this because the day's accrual update is hapening in Start Processing part of the function. This function is simply creating the accrual file
                                'Append daily accrual details in the Daily Accrual table
                                'If CVDate(xAsOf) <= CVDate(Date) Then
                                '    'Accumulate daily accruals of the employee from begining of the entitlement period to today's date and update ED_VAC
                                '    dblEntitleUpd# = Round(dblEntitle# + dblNewDailyEnt, 4)
                                '
                                '    'Append the Daily Accrual table and update with Process Date as this will append to employee's Vacation Entitlement
                                '    Call Append_Daily_Accrul_File(rsHREmp("ED_EMPNBR"), clpCode(0).Text, clpCode(1).Text, clpPT.Text, clpCode(2).Text, dlpDateRange(0).Text, dlpDateRange(1).Text, dblNewEntitle, xAsOf, dblNewDailyEnt, Date, Round(Accrued_ToDate(rsHREmp("ED_EMPNBR"), clpCode(0).Text, clpCode(1).Text, clpPT.Text, clpCode(2).Text, dlpDateRange(0).Text, dlpDateRange(1).Text, xAsOf) + dblNewDailyEnt, 4), False)
                                '
                                '    'Update ED_VAC in HREMP table with day's accrual (from begining of the entitlement period to today's date)
                                '    rsHREmp("ED_VAC") = dblEntitleUpd
                                '
                                '    'Append in Accrual table as well
                                '    xComments = "Current Vac. Ent. Chg from " & dblEntitle# & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(rsHREmp("ED_PVAC")), 0, rsHREmp("ED_PVAC")) + IIf(IsNull(rsHREmp("ED_VAC")), 0, rsHREmp("ED_VAC"))) - IIf(IsNull(rsHREmp("ED_VACT")), 0, rsHREmp("ED_VACT"))
                                '    Call Append_Accrual(rsHREmp("ED_EMPNBR"), "VAC", CVDate(Format(Now, "mm/dd/yyyy")), dblNewDailyEnt, "U", xComments)
                                '
                                '    'Update local variable with the new ED_VAC so it can be used for next day for accumulation (above)
                                '    dblEntitle# = rsHREmp("ED_VAC")
                                '    rsHREmp.Update
                                'Else
                                    'Employee should accrue to the max on the Last Accrual Day. This means if their Annual Accrual is 200 and their last day's accrual is setting them
                                    'to 199.9835 as Accrued to Date, then that Day's Accrual should be rounded to make Accrued to Date as 200. Or if Annual Accrual is 160 but their
                                    'last day's Accrual is setting them to 160.016 as Accrued to Date, then that Day's Accural should round down to make Accrued to Date as 160.
                                    'And also the employee should not have moved from one level to another.
                                    'And also the employee should not have started earning mid year due to the Start Date of the employment to earn daily accrual
                                    If CVDate(xAsOf) = CVDate(xToDate) And flgLvlChanged = False And flgMidYearStart = False Then
                                        'Accrual to Date as of day before last day
                                        xDayB4AccToDate = Round(Accrued_ToDate(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), xFromDate, xToDate, DateAdd("d", -1, CVDate(xAsOf))), 4)
                                        
                                        'Last Day's Accrual
                                        xLstAccToDate = xDayB4AccToDate + dblNewDailyEnt
                                        
                                        'Annual Accrual > Accrued to Date
                                        If Round(dblNewEntitle, 4) > Round(xLstAccToDate, 4) Then
                                            'Round Up the Daily Accrual to Annual Accrual, e.g. 200 > 199.9835
                                            'Get the difference between Day Before's Accrual To Date and Annual Accrual that will be the Last Day's Daily Accrual
                                            dblNewDailyEnt = dblNewEntitle - xDayB4AccToDate
                                            
                                        ElseIf Round(dblNewEntitle, 4) < Round(xLstAccToDate, 4) Then
                                            'Round Down the Daily Accrual to Annual Accrual, e.g. 160 < 160.16
                                            'Get the difference between Day Before's Accrual To Date and Annual Accrual that will be the Last Day's Daily Accrual
                                            dblNewDailyEnt = dblNewEntitle - xDayB4AccToDate
                                        Else
                                            'Don't do anything as Annual Accrual = Accrued to Date
                                        End If
                                    End If
                                    
                                    'Future day's accrual
                                    'Add the daily accrual to the Daily Accrual details table, not Processed yet
                                    Call Append_Daily_Accrul_File(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), xEffDate, rsDailyVacEntlmt("VD_FRDATE"), rsDailyVacEntlmt("VD_TODATE"), dblNewEntitle, xAsOf, dblNewDailyEnt, "", Round(Accrued_ToDate(rsHREmp("ED_EMPNBR"), rsDailyVacEntlmt("VD_ORG"), rsDailyVacEntlmt("VD_EMP"), rsDailyVacEntlmt("VD_PT"), rsDailyVacEntlmt("VD_EMPEXCL"), rsDailyVacEntlmt("VD_FRDATE"), rsDailyVacEntlmt("VD_TODATE"), xAsOf) + dblNewDailyEnt, 4), False)
                                'End If
                                
                            End If
                                        
lblNextDay:
                            'Next day of the Vacation Entitlement Start Date
                            xAsOf = DateAdd("d", 1, CVDate(xAsOf))
                        Loop
lblNextRec:
                        If rsJOB.State <> 0 Then rsJOB.Close
                        Set rsJOB = Nothing
                        
                        rsHREmp.MoveNext
                        DoEvents
                    Wend
                End If
                rsDailyVacEntlmt.Close
                Set rsDailyVacEntlmt = Nothing
                
                Create_Daily_Accrual_File = True
                
                MDIMain.panHelp(0).FloodType = 0
            End If
Next_Rule:
            rsHREmp.Close
            Set rsHREmp = Nothing

            rsDailyVacEnt.MoveNext
        Loop
    End If
    rsDailyVacEnt.Close
    Set rsDailyVacEnt = Nothing
    
    Create_Daily_Accrual_File = True
End Function

