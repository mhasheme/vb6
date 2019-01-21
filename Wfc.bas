Attribute VB_Name = "WFC"
'Created by Jaddy 7/22/99
Public Const glbDiscipStartDate = "Mar 09,2003" '"Jan 1,2004" '"Mar 09,2003"
Public Const glbWFCNameChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz ‘-~." 'Ticket #14154
Global glbEmpPlant
Global glbWFCEmailTest As Boolean
Global glbWFCTermEmail
Global glbWFCHrsSal As Boolean
Global glbBenefitAccount
Global xBreakInService
Global xLeaveMonth
Global xExpYears
Global xEmpExpYears
Global x15DayLogic As Boolean
Global glbNGS_OnFlag As Boolean
Global glbWFC_US_Ben_Trans As Boolean 'Ticket #23247 Franks 04/22/2013
Global glbWFCNGSSubGroup As String
Global glbWFCPayGroup As String
Global glbEmpDiv As String
'Ticket #19954 Pen Audit - begin
Global UpdPenAudit As Boolean
Global UpdPenAudDirect As Boolean
Global toSOURCE
Global toTYPE
Global oDBStatus
Global oDBStaDate
Global oEntryDate
Global oExitDate
Global oEarnPen
Global oContSrv
Global oCredSrv
Global oBenRate
Global glbHRSoftType As String
Global glbHRSoftAction As String
'Ticket #19954 Pen Audit - end
'Ticket #20885 Franks 12/01/2011 for Samuel - begin
Global glbEmpAdminBy As String
Global glbEmpSection As String
Global glbEmpRegion As String
'Ticket #20885 Franks 12/01/2011 for Samuel - end
Global locWFCEmpID
Global glbDivTranInPlant As String 'Ticket #25221 Franks 03/17/2014
Global glbWFCNewPosJob 'Ticket #27827 Franks 12/02/2015
Global glbWFCNewPosDiv 'Ticket #27827 Franks 12/02/2015
Global glbWFCNewPosStatus 'Ticket #27827 Franks 12/02/2015
Global glbWFC_IncentivePlanFlag As Boolean 'Ticket #29013 Franks 08/23/2016
Global glbWFC_IncePlanID As Long ' Integer 'Ticket #29013 Franks 08/25/2016
Global glbWFC_IPPopFormName As String  'Ticket #29013 Franks 08/29/2016
Global glbWFC_CancelTransaction As Boolean
Global glbWFCContractEmployee As Boolean

Sub Band_AddItem(frmName As Form)
Dim rsSBand As New ADODB.Recordset
Dim SQLQ As String, xItem As String
  frmName.cmbBand.AddItem "" 'Ticket# 6467 07/06/2004 Frank, Allow user put blank
  frmName.cmbBand.AddItem "A"
  frmName.cmbBand.AddItem "B"
  frmName.cmbBand.AddItem "C"
  frmName.cmbBand.AddItem "D"
  frmName.cmbBand.AddItem "E"
  frmName.cmbBand.AddItem "F"
  frmName.cmbBand.AddItem "G"
  frmName.cmbBand.AddItem "H"
  frmName.cmbBand.AddItem "I"
  frmName.cmbBand.AddItem "J"
  'Add by Franks on Jul 11,02 for ticket #2546
  xItem = ",A,B,C,D,E,F,G,H,I,J,"
  SQLQ = "SELECT [BAND] AS GBAND FROM WFC_Salary_Administration GROUP BY [BAND]"
  rsSBand.Open SQLQ, gdbAdoIhrWFC, adOpenStatic
  Do While Not rsSBand.EOF
    If InStr(xItem, Trim(rsSBand("GBAND")) & ",") = 0 Then
        frmName.cmbBand.AddItem Trim(rsSBand("GBAND"))
    End If
    rsSBand.MoveNext
  Loop
  rsSBand.Close
  'Add by Franks on Jul 11,02 for ticket #2546
End Sub

Sub Band_Setup(frmName As Form)
  frmName.lblBand.Visible = glbWFC
  frmName.cmbBand.Visible = glbWFC
  Band_AddItem frmName
End Sub

Sub cmbBand_SETUP(frmName As Form)
Dim X%, Temp
frmName.txtBand.Visible = (Trim(frmName.txtBand) = "")
'frmName.cmbBand.ListIndex = 0
Temp = False
For X% = 0 To frmName.cmbBand.ListCount - 1
  If frmName.cmbBand.List(X%) = frmName.txtBand Then
    frmName.cmbBand.ListIndex = X%
    Temp = True
    Exit For
  End If
Next
If Temp = False Then
  frmName.txtBand.Visible = True
  If frmName.Caption = "Salary Grids" Then
    frmName.cmbBand = frmName.txtBand
  End If
End If
End Sub

Sub cmbCurrencyIndicator_setup(frmName As Form)
Dim X%
frmName.txtCurrencyIndicator.Visible = (Trim(frmName.txtCurrencyIndicator) = "")
'frmName.cmbCurrencyIndicator.ListIndex = 0
For X% = 0 To 1
  If frmName.cmbCurrencyIndicator.List(X%) = Left(frmName.txtMarketLine, 2) Then
    frmName.cmbCurrencyIndicator.ListIndex = X%
    Exit For
  End If
Next
End Sub

Sub setMarketLine(frmName As Form)
Dim X%, Temp
frmName.txtMarketLine.Visible = (Trim(frmName.txtMarketLine) = "")
'frmName.cmbMarketLine.ListIndex = 0
Temp = False
For X% = 0 To frmName.cmbMarketLine.ListCount - 1
  If Left(frmName.cmbMarketLine.List(X%), 4) = frmName.txtMarketLine Then
    frmName.cmbMarketLine.ListIndex = X%
    Temp = True
    Exit For
  End If
Next
If frmName.Caption = "Salary Grids" Then
    If Temp = False Then
      frmName.cmbMarketLine = frmName.txtMarketLine
    End If
End If
End Sub

Sub CurrencyIndicator_AddItem(frmName As Form)
  frmName.cmbCurrencyIndicator.AddItem "CA"
  frmName.cmbCurrencyIndicator.AddItem "US"
End Sub



Sub MarketLine_Desc(frmName As Form)
Dim Temp
If frmName.txtMarketLine.Visible Then
  Temp = frmName.txtMarketLine
Else
  Temp = frmName.cmbMarketLine
End If
  Select Case Temp
  Case "CA01": frmName.lblMLine = "MISSISSAUGA-Band j"
  Case "CA02": frmName.lblMLine = "MISSISSAUGA-Legal"
  Case "CA03": frmName.lblMLine = "MISSISSAUGA-IT"
  Case "CA04": frmName.lblMLine = "MISSISSAUGA,Kipling"
  Case "CA05": frmName.lblMLine = "Kipling-Salary Option 1"
  Case "CA06": frmName.lblMLine = "Kipling-Salary Option 3"
  Case "CA07": frmName.lblMLine = "Kipling-Salary Option 4"
  Case "CA08": frmName.lblMLine = "St.Jerome"
  Case "CA09": frmName.lblMLine = "St.Jerome-Salary Option 1"
  Case "CA10": frmName.lblMLine = "Whitby, Tilbury, Sarnia"
  Case "CA11": frmName.lblMLine = "Whitby, Tilbury, Sarnia"
  Case "CA12": frmName.lblMLine = "Tibury-Salary Option 3"
  Case "US01": frmName.lblMLine = "Troy, Whitmore Lake, Chesterfield, Romulus"
  Case "US02": frmName.lblMLine = "Troy-Band J"
  Case "US03": frmName.lblMLine = "Chesterfield-Salary Option 3"
  Case "US04": frmName.lblMLine = "Addison, Fairless Hills"
  Case "US05": frmName.lblMLine = "Fairless Hills-Salary Option 2"
  Case "US06": frmName.lblMLine = "Atlanta"
  Case "US07": frmName.lblMLine = "Brodhead"
  Case "US08": frmName.lblMLine = "Brodhead-Salary Option 3"
  Case "US09": frmName.lblMLine = "Chattanooga"
  Case "US10": frmName.lblMLine = "Chattanooga-Salary Option 1"
  Case "US11": frmName.lblMLine = "Fremont, Columbus"
  Case "US12": frmName.lblMLine = "Kansas City, Riverside, Wdbg Sqcg Cntr."
  Case "US13": frmName.lblMLine = "Riverside/Wdbg Sqcg-Salary Option 1"
  Case "US14": frmName.lblMLine = "Riverside/Wdbg Sqcg-Salary Option 3"
  Case "US15": frmName.lblMLine = "St. Peters"
  Case "US16": frmName.lblMLine = "St. Peters-Salary Option 2"
  Case Else: frmName.lblMLine = ""
  End Select
End Sub

Sub lblBand_Setup(frmName As Form)
  frmName.lblBand.Visible = glbWFC
  frmName.lblBANDCode.Visible = glbWFC
End Sub

Sub MarketLine_AddItem(frmName As Form)
Dim rsSMarketLind As New ADODB.Recordset
Dim SQLQ As String, xItem As String
  If Not glbWFC Then Exit Sub
  frmName.cmbMarketLine.AddItem "CA01"
  frmName.cmbMarketLine.AddItem "CA02"
  frmName.cmbMarketLine.AddItem "CA03"
  frmName.cmbMarketLine.AddItem "CA04"
  frmName.cmbMarketLine.AddItem "CA05"
  frmName.cmbMarketLine.AddItem "CA06"
  frmName.cmbMarketLine.AddItem "CA07"
  frmName.cmbMarketLine.AddItem "CA08"
  frmName.cmbMarketLine.AddItem "CA09"
  frmName.cmbMarketLine.AddItem "CA10"
  frmName.cmbMarketLine.AddItem "CA11"
  frmName.cmbMarketLine.AddItem "CA12"
  xItem = ",CA01,CA02,CA03,CA04,CA05,CA06,CA07,CA08,CA09,CA10,CA11,CA12,"
  SQLQ = "SELECT [MarketLine] AS GMarketLine FROM WFC_Salary_Administration "
  SQLQ = SQLQ & "WHERE LEFT(MarketLine,2) = 'CA' GROUP BY [MarketLine]"
  rsSMarketLind.Open SQLQ, gdbAdoIhrWFC, adOpenStatic
  Do While Not rsSMarketLind.EOF
    If InStr(xItem, Trim(rsSMarketLind("GMarketLine")) & ",") = 0 Then
        frmName.cmbMarketLine.AddItem Trim(rsSMarketLind("GMarketLine"))
    End If
    rsSMarketLind.MoveNext
  Loop
  rsSMarketLind.Close
  frmName.cmbMarketLine.AddItem "US01"
  frmName.cmbMarketLine.AddItem "US02"
  frmName.cmbMarketLine.AddItem "US03"
  frmName.cmbMarketLine.AddItem "US04"
  frmName.cmbMarketLine.AddItem "US05"
  frmName.cmbMarketLine.AddItem "US06"
  frmName.cmbMarketLine.AddItem "US07"
  frmName.cmbMarketLine.AddItem "US08"
  frmName.cmbMarketLine.AddItem "US09"
  frmName.cmbMarketLine.AddItem "US10"
  frmName.cmbMarketLine.AddItem "US11"
  frmName.cmbMarketLine.AddItem "US12"
  frmName.cmbMarketLine.AddItem "US13"
  frmName.cmbMarketLine.AddItem "US14"
  frmName.cmbMarketLine.AddItem "US15"
  frmName.cmbMarketLine.AddItem "US16"
  frmName.cmbMarketLine.AddItem "US17"
  xItem = "US01,US02,US03,US04,US05,US06,US07,US08,"
  xItem = xItem & "US09,US10,US11,US12,US13,US14,US15,US16,US17,"
  SQLQ = "SELECT [MarketLine] AS GMarketLine FROM WFC_Salary_Administration "
  SQLQ = SQLQ & "WHERE LEFT(MarketLine,2) = 'US' GROUP BY [MarketLine]"
  rsSMarketLind.Open SQLQ, gdbAdoIhrWFC, adOpenStatic
  Do While Not rsSMarketLind.EOF
    If InStr(xItem, Trim(rsSMarketLind("GMarketLine")) & ",") = 0 Then
        frmName.cmbMarketLine.AddItem Trim(rsSMarketLind("GMarketLine"))
    End If
    rsSMarketLind.MoveNext
  Loop
  rsSMarketLind.Close

End Sub

Sub MARKETLINE_Setup(frmName As Form)
  Dim X%
  frmName.lblMarketLine.Visible = glbWFC
  frmName.cmbMarketLine.Visible = glbWFC
  frmName.lblMLine.Visible = glbWFC
  For X% = 0 To 2
    frmName.lblsalstate(X%).Visible = glbWFC
  Next
'  frmName.lblTitle(13).Visible = glbWFC  'Jaddy 10/16/99
  If glbWFC Then
    MarketLine_AddItem frmName
  End If
End Sub

Public Sub Whitby60daysRule(xEmpNo, xType)
Dim rsTemp As New ADODB.Recordset
Dim rsTem2 As New ADODB.Recordset
Dim rsTemEmp As New ADODB.Recordset
Dim SQLQ, xCodeList, xNextDiscipStep
Dim xDaysDiff, xSuperID, xDiscipStep
Dim xTflag As Boolean, xDiscipCode, I
Dim xDcode1, xDcode2, xDdate1, xDdate2

    'Enable this function Ticket# 6656
    'Disable it until Whitby is ready
    'Exit Sub
    
'After 60 days of no uncontrolled absence, the system can subtract a point from the employee.
'If the employee qualifies for the reduction in points, an attendance record is created for that day.
'The attendance reason will be Point Reduction (code PR) and the point value will be -1.
'Check both Absent Reason Codes and PR, get the recent date to compare with today

'Re-setup the Next Disciplinary Step, also it can work for First time Recalaulate
'Check both Attendance and HR_COUNSEL tables, find the latest date of Disciplinary action
'to set Next Disciplinary Step

    'If xEmpNo = "ALL" Then '"RECAL" Then
        Screen.MousePointer = HOURGLASS
        MDIMain.panHelp(0).FloodType = 1
        'If No Disciplianry attendance record, get the current step from Counsel table
        'Get the Absent reason codes
        xCodeList = "('***'"
        SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsTemp.EOF
            xCodeList = xCodeList & ",'" & rsTemp("DS_DISCIPLINE") & "'"
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        xCodeList = xCodeList & ",'PR')"
        
        SQLQ = "SELECT ED_EMPNBR,ED_DISCIPLINENEXT FROM HREMP WHERE ED_SECTION='WHBY' " '& glbSeleDeptUn
        If Not xEmpNo = "ALL" Then
            'For single employee
            SQLQ = SQLQ & "AND ED_EMPNBR = " & xEmpNo
        End If
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        I = 0
        Do While Not rsTemp.EOF
            MDIMain.panHelp(0).FloodPercent = (I / rsTemp.RecordCount) * 100: I = I + 1

            'If IsNull(rsTemp("ED_DISCIPLINENEXT")) Then
                xDcode1 = "": xDcode2 = ""
                xDdate1 = "": xDdate2 = ""
                xTflag = False
                'Check if there is Disciplianry attendance record
                SQLQ = "SELECT AD_EMPNBR,AD_DISCIPLINE,AD_DOA FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & rsTemp("ED_EMPNBR") & " "
                SQLQ = SQLQ & "AND AD_DISCIPLINE IN " & xCodeList & " "
                SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
                SQLQ = SQLQ & "ORDER BY AD_EMPNBR,AD_DOA DESC "
                rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If rsTem2.EOF Then
                    xTflag = True
                Else
                    xDcode1 = rsTem2("AD_DISCIPLINE")
                    xDdate1 = rsTem2("AD_DOA")
                End If
                rsTem2.Close
 
                SQLQ = "SELECT CL_EMPNBR,CL_TYPE,CL_INCDATE FROM HR_COUNSEL WHERE CL_EMPNBR = " & rsTemp("ED_EMPNBR") & " "
                SQLQ = SQLQ & "AND NOT(CL_INCDATE IS NULL) "
                SQLQ = SQLQ & "AND CL_REASON='ATT' "
                SQLQ = SQLQ & "AND CL_TYPE IN " & xCodeList & " ORDER BY CL_INCDATE DESC " 'CL_TYPE DESC"
                rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                xDiscipCode = ""
                If Not rsTem2.EOF Then
                    xDiscipCode = rsTem2("CL_TYPE")
                    xDcode2 = rsTem2("CL_TYPE")
                    If IsDate(rsTem2("CL_INCDATE")) Then
                        xDdate2 = rsTem2("CL_INCDATE")
                    End If
                Else
                    xTflag = True
                End If
                rsTem2.Close
                
                If Not xTflag Then
                    If Not (Len(xDdate1) = 0 And Len(xDdate2) = 0) Then
                        If IsDate(xDdate1) Then
                            xDiscipCode = xDcode1
                        End If
                        If IsDate(xDdate2) Then
                            xDiscipCode = xDcode2
                        End If
                        If IsDate(xDdate1) And IsDate(xDdate2) Then
                            If CVDate(xDdate1) = CVDate(xDdate2) Then
                                If Val(xDcode2) >= Val(xDcode1) Then
                                    xDiscipCode = xDcode2
                                Else
                                    xDiscipCode = xDcode1
                                End If
                            Else
                                If CVDate(xDdate1) >= CVDate(xDdate2) Then
                                    xDiscipCode = xDcode1
                                Else
                                    xDiscipCode = xDcode2
                                End If
                            End If
                        End If
                    End If
                    
                    '-111
                    'If Len(xDiscipCode) > 0 Then
                    '    xDiscipStep = 0
                    '    'Get the Step value
                    '    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_DISCIPLINE = '" & xDiscipCode & "' "
                    '    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    '    If Not rsTem2.EOF Then
                    '        If rsTem2("DS_STEPNO") > 0 Then
                    '            rsTemp("ED_DISCIPLINENEXT") = rsTem2("DS_STEPNO") + 1
                    '            rsTemp.Update
                    '        End If
                    '    End If
                    '    rsTem2.Close
                    'End If
                End If
                
                '-111
                If Len(xDiscipCode) > 0 Then
                    xDiscipStep = 0
                    'Get the Step value
                    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_DISCIPLINE = '" & xDiscipCode & "' "
                    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not rsTem2.EOF Then
                        If rsTem2("DS_STEPNO") > 0 Then
                            xDiscipStep = rsTem2("DS_STEPNO") + 1
                            rsTemp("ED_DISCIPLINENEXT") = rsTem2("DS_STEPNO") + 1
                            rsTemp.Update
                        End If
                    End If
                    rsTem2.Close
                End If
                '-111
                
                'End If
            'End If
            rsTemp.MoveNext
        Loop
        rsTemp.Close
        MDIMain.panHelp(0).FloodType = 0
        Screen.MousePointer = DEFAULT
        'Exit Sub
    'End If
    
    
    
    '=======================================================
    'Don't check 60days rule When delete Attendance record
    If xType = "D" Then Exit Sub
    
    '60 days rule Calculation
    'For all Whitby employees
    SQLQ = "SELECT ED_EMPNBR,ED_DISCIPLINENEXT FROM HREMP WHERE ED_SECTION='WHBY' "
    If Not xEmpNo = "ALL" Then
        'For single employee
        SQLQ = SQLQ & "AND ED_EMPNBR = " & xEmpNo
    End If
    

    If rsTemEmp.State <> 0 Then rsTemEmp.Close
    rsTemEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    
    'Get the Absent reason codes
    xCodeList = "('***'"
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'ADRE' AND TB_ABSENCE <> 0 AND TB_USR2>0"
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsTemp.EOF
        xCodeList = xCodeList & ",'" & rsTemp("TB_KEY") & "'"
        rsTemp.MoveNext
    Loop
    rsTemp.Close
    xCodeList = xCodeList & ",'" & "PR" & "'" & ")"
        
    I = 0
    Do While Not rsTemEmp.EOF
        MDIMain.panHelp(0).FloodPercent = (I / rsTemEmp.RecordCount) * 100: I = I + 1
        xEmpNo = rsTemEmp("ED_EMPNBR")
        'Find the value of  the Next Disciplinary Step
        'If it is null or 1, don't do anything
        xNextDiscipStep = 1
        
        If Not IsNull(rsTemEmp("ED_DISCIPLINENEXT")) Then
            xNextDiscipStep = rsTemEmp("ED_DISCIPLINENEXT")
        End If
        
        If xNextDiscipStep = 1 Then
            GoTo Next_Lin
            'Exit Sub
        End If

        'Check Attendance record
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND AD_REASON IN " & xCodeList & " "
        SQLQ = SQLQ & "AND AD_DOA >= " & Date_SQL(CVDate(glbDiscipStartDate)) & " "
        SQLQ = SQLQ & "ORDER BY AD_EMPNBR,AD_DOA DESC "
        
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsTemp.EOF Then
            rsTemp.Close
            GoTo Next_Lin
        End If
        xDaysDiff = DateDiff("d", rsTemp("AD_DOA"), Date)
        If xDaysDiff <= 60 Then
            rsTemp.Close
            GoTo Next_Lin
        End If
        If Not IsNull(rsTemp("AD_SUPER")) Then
            xSuperID = rsTemp("AD_SUPER")
        Else
            xSuperID = 0
        End If

        
        'the Next Disciplinary Step -1
        xDiscipStep = 0
        SQLQ = "SELECT ED_EMPNBR, ED_DISCIPLINENEXT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
        If rsTem2.State <> 0 Then rsTem2.Close
        rsTem2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsTem2.EOF Then
            If Not IsNull(rsTem2("ED_DISCIPLINENEXT")) Then
                If rsTem2("ED_DISCIPLINENEXT") > 1 Then
                    xDiscipStep = rsTem2("ED_DISCIPLINENEXT") - 1 '-111 rsTemp(
                    xDiscipStep = WhitbyPreStep(xEmpNo, rsTem2("ED_DISCIPLINENEXT") - 2)
                    rsTem2("ED_DISCIPLINENEXT") = xDiscipStep + 1
                    rsTem2.Update
                End If
            End If
        End If
        rsTem2.Close
        
        'Get current Discipline code
        xDiscipCode = ""
        SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_STEPNO = " & xDiscipStep & " "
        rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTem2.EOF Then
            xDiscipCode = rsTem2("DS_DISCIPLINE")
        End If
        rsTem2.Close
            
        'Add PR record for this employee
        rsTemp.AddNew
        rsTemp("AD_COMPNO") = "001"
        rsTemp("AD_EMPNBR") = xEmpNo
        rsTemp("AD_DOA") = Date
        rsTemp("AD_REASON") = "PR"
        rsTemp("AD_HRS") = 0
        rsTemp("AD_POINT") = -1
        If Len(xDiscipCode) > 0 Then
            rsTemp("AD_DISCIPLINE") = xDiscipCode
        End If
        rsTemp("AD_LDATE") = Date
        rsTemp("AD_LTIME") = Time$
        rsTemp("AD_LUSER") = glbUserID
        If xSuperID > 0 Then rsTemp("AD_SUPER") = xSuperID
        rsTemp.Update
        rsTemp.Close
        
        'If xDiscipStep > 0 Then
        '    'Reset the current Disciplinary
        '    'To false
        '    SQLQ = "UPDATE HR_COUNSEL SET CL_COMPLETED = 0 WHERE CL_EMPNBR = " & xEmpNo & " "
        '    SQLQ = SQLQ & "AND CL_COMPLETED <> 0 "
        '    gdbAdoIhr001.Execute SQLQ
        '    'Get current Disciplinary Code
        '    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_STEPNO = " & xDiscipStep - 1 & " "
        '    rsTem2.Open SQLQ, gdbAdoIhr001, adOpenStatic
        '    If Not rsTem2.EOF Then
        '        SQLQ = "UPDATE HR_COUNSEL SET CL_COMPLETED = -1 WHERE CL_EMPNBR = " & xEmpNo & " "
        '        SQLQ = SQLQ & "AND CL_TYPE = '" & rsTem2("DS_DISCIPLINE") & "' "
        '        gdbAdoIhr001.Execute SQLQ
        '    End If
        '    rsTem2.Close
        'End If
Next_Lin:
        rsTemEmp.MoveNext
    Loop
    rsTemEmp.Close
    MDIMain.panHelp(0).FloodType = 0
    Screen.MousePointer = DEFAULT
    
End Sub

Public Function WhitbyPreStep(xEmpNo, xStep)
Dim rsDisci As New ADODB.Recordset
Dim rsCounsel As New ADODB.Recordset
Dim SQLQ, xPreStep, I, xMum
'Dim xArray(10, 2)
    xPreStep = xStep
    SQLQ = "SELECT * FROM HR_DISCIPLINE_STEPS WHERE DS_STEPNO <= " & xStep & " ORDER BY DS_STEPNO DESC"
    rsDisci.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsDisci.EOF Then
        rsDisci.Close
        GoTo end_line
    End If
    xPreStep = 1
    Do While Not rsDisci.EOF
        SQLQ = "SELECT CL_EMPNBR FROM HR_COUNSEL WHERE CL_TYPE = '" & rsDisci("DS_DISCIPLINE") & "' "
        SQLQ = SQLQ & "AND  CL_EMPNBR = " & xEmpNo
        If rsCounsel.State <> 0 Then rsCounsel.Close
        rsCounsel.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsCounsel.EOF Then
            xPreStep = rsDisci("DS_STEPNO")
            rsCounsel.Close
            rsDisci.Close
            GoTo end_line
        End If
        rsCounsel.Close
        rsDisci.MoveNext
    Loop
    
end_line:
    WhitbyPreStep = xPreStep
End Function

Public Function GetBenCertificateNo(xBenGroup, xEmpNo)
Dim rsBenGrpMatrix As New ADODB.Recordset
Dim rsEmpPayID As New ADODB.Recordset
Dim SQLQ As String, xPrefix As String, xDIV As String, xPayID As String
Dim xCerNo As String
    xCerNo = "": xPrefix = "": xDIV = "": xPayID = ""
    glbBenefitAccount = ""
    SQLQ = "SELECT ED_DIV,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmpPayID.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpPayID.EOF Then
        xPayID = rsEmpPayID("ED_PAYROLL_ID")
        xDIV = rsEmpPayID("ED_DIV")
    End If
    rsEmpPayID.Close
    If Len(xPayID) > 0 And Len(xDIV) > 0 Then
        SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE BM_BENEFIT_GROUP = '" & xBenGroup & "' "
        SQLQ = SQLQ & "AND BM_DIV = '" & xDIV & "' "
        rsBenGrpMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsBenGrpMatrix.EOF Then
            xPrefix = IIf(IsNull(rsBenGrpMatrix("BM_CERTIFICATE_PREFIX")) Or rsBenGrpMatrix("BM_CERTIFICATE_PREFIX") = "", "", rsBenGrpMatrix("BM_CERTIFICATE_PREFIX"))
            glbBenefitAccount = rsBenGrpMatrix("BM_BENEFIT_ACCOUNT")
        End If
        rsBenGrpMatrix.Close
        If Len(xPrefix) > 0 Then
            xCerNo = Trim(xPrefix) & Trim(xPayID)
            xCerNo = Right("000000000000" & xCerNo, 12)
        End If
    End If
    GetBenCertificateNo = xCerNo
End Function

Public Function InvalidCharInStr(Str As String, Filter As String) As String
    Dim I As Integer, buf As String, CHAR As String * 1
    buf = ""
    For I = 1 To Len(Str)
        CHAR = Mid(Str, I, 1)
        If InStr(Filter, CHAR) = 0 Then buf = CHAR
    Next I
    InvalidCharInStr = buf
End Function

Public Sub WFCCNDBeneAuditFlag(Optional empNo)
'Ticket #15818
'if benefit updates Audit, make update flag = "Y" (Canadian Employees only)
Dim SQLQ As String
    SQLQ = "UPDATE HRAUDIT SET AU_UPLOAD = 'Y' "
    SQLQ = SQLQ & "WHERE AU_BCODE_TABL = 'BNCD' AND (AU_BCODE IS NOT NULL ) "
    SQLQ = SQLQ & "AND (AU_UPLOAD IS NULL OR AU_UPLOAD = 'N') "
    If Not IsMissing(empNo) Then
        SQLQ = SQLQ & "AND AU_EMPNBR = " & empNo & " "
    End If
    SQLQ = SQLQ & "AND (AU_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_COUNTRY = 'CANADA') OR (AU_EMPNBR IN (SELECT ED_EMPNBR FROM TERM_HREMP WHERE ED_COUNTRY = 'CANADA')))"
    gdbAdoIhr001.Execute SQLQ
End Sub

Public Sub WFCPensionBeneficiary(xEmpNo, xBenCode, Optional xTermSEQ = 0, Optional NewUpt = "NEW", Optional xBensChg, Optional xCurrentBens, Optional xAsOfBens, Optional xTransPenType, Optional xBenBame)
'see Pension Phase II.doc in W:\2008 Projects\Pension\Pension Phase II folder
Dim rsEmp As New ADODB.Recordset
Dim rsHRBenF As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim SQLQ As String
Dim xSection, xPenType, xDBStatus, xSalHly
Dim xDOB, xEarlyRet, xNorRet, xLateRet, xDATE, xYear, xSIN, xUnion
Dim xUpdFlag As String
Dim xAskSIN As Boolean
Dim xLocSpoSIN
Dim xOrgCurrent As Integer

    'Ticket #22776 Franks 11/06/2012 - check both active and terminated employees
    If xTermSEQ > 0 Then
        SQLQ = "SELECT * FROM Term_HREMP WHERE TERM_SEQ = " & xTermSEQ & " "
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Else
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    End If
    'Ticket #22776 Franks 11/06/2012 - end
    If Not rsEmp.EOF Then
        If IsNull(rsEmp("ED_EMPTYPE")) Then Exit Sub
        If Not rsEmp("ED_EMPTYPE") = "Y" Then Exit Sub
        If IsNull(rsEmp("ED_SIN")) Then Exit Sub
        If IsNull(rsEmp("ED_SECTION")) Then Exit Sub
        If IsNull(rsEmp("ED_DOH")) Then Exit Sub
        If IsNull(rsEmp("ED_DOB")) Then Exit Sub
        If IsNull(rsEmp("ED_ORG")) Then Exit Sub
        xSIN = rsEmp("ED_SIN")
        xSection = rsEmp("ED_SECTION")
        xUnion = rsEmp("ED_ORG")
        xYear = Year(rsEmp("ED_DOH"))
        If Not IsMissing(xTransPenType) Then 'Ticket #24275 Franks 08/27/2013
            If Len(xTransPenType) > 0 Then
                xPenType = xTransPenType
            Else
                xPenType = getDBType(xSection, xUnion, "PenType", rsEmp("ED_DOH")) 'Ticket #26707 Franks 02/25/2015
            End If
        Else
            xPenType = getDBType(xSection, xUnion, "PenType", rsEmp("ED_DOH")) 'Ticket #26707 Franks 02/25/2015
        End If
        xSalHly = getDBType(xSection, xUnion, "HlySal")
        If Len(xPenType) = 0 Then
            Exit Sub
        End If
        If Len(xSalHly) = 0 Then
            Exit Sub
        End If

        If Not IsNull(rsEmp("ED_DOB")) Then
            xDOB = rsEmp("ED_DOB")
        Else
            xDOB = Date
        End If
        
        'rsHRBenF
        If xTermSEQ > 0 Then
            SQLQ = "SELECT * FROM Term_HRBENS "
            SQLQ = SQLQ & " WHERE (1=1) " 'BD_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & " AND TERM_SEQ = " & xTermSEQ & " "
        Else
            SQLQ = "SELECT * FROM HRBENS "
            SQLQ = SQLQ & " WHERE BD_EMPNBR = " & xEmpNo & " "

        End If
        SQLQ = SQLQ & " AND BD_BCODE = '" & xBenCode & "' "
        If Not IsMissing(xBenBame) Then
            SQLQ = SQLQ & " AND BD_BNAME = '" & xBenBame & "' "
        End If
        If Not IsMissing(xTransPenType) Then 'Ticket #24337 Franks 09/10/2013
            QLQ = SQLQ & " AND BD_PENSIONTYPE = '" & xTransPenType & "' "
        End If
        SQLQ = SQLQ & " ORDER BY BD_BCODE, BD_LDATE "
        rsHRBenF.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        
        'Ticket #21021 Franks 02/28/2012
        'If Not rsHRBenF.EOF Then
        xAskSIN = False
        xLocSpoSIN = ""
        If Not rsHRBenF.EOF Then
            If rsHRBenF.RecordCount = 1 Then 'only ask for the fist time
                xAskSIN = True
            Else
                'get Spouse SIN
                SQLQ = "SELECT * FROM HRP_PENSION_BENEFICIARY WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' "
                SQLQ = SQLQ & "AND NOT (PE_SPOUSE_SIN IS NULL OR PE_SPOUSE_SIN = '') "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsPen.EOF Then
                    xLocSpoSIN = rsPen("PE_SPOUSE_SIN")
                End If
                rsPen.Close
            End If
        End If
        Do While Not rsHRBenF.EOF
            'add beneficiary master record - begin
            SQLQ = "SELECT * FROM HRP_PENSION_BENEFICIARY WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' "  '"' AND PE_HRLYSAL = '" & xSalHly & "' "
            If NewUpt = "NEW" Then 'Ticket #22884 'Ticket #23435 Frank 03/19/2013 - added this back
                SQLQ = SQLQ & "AND PE_BEN_NAME = '" & rsHRBenF("BD_BNAME") & "' " 'Ticket #21021 Franks 02/28/2012 for Separation Agreement
            End If
            'Else 'PE_SPE_DATE
            '    SQLQ = SQLQ & "AND PE_SPE_DATE IS NULL " 'only update the record without PE_SPE_DATE
            'End If
            
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            xUpdFlag = ""
            'If rsPen.EOF Then
            If rsPen.EOF Then
                xUpdFlag = "NEW" '
            End If
            If rsPen.EOF And xAskSIN Then
                'Ticket #21543 Franks 02/07/2012, Jerry & MZ asked to remove this logic
                ''If UCase(rsHRBenF("BD_RELATE")) = "SPOUSE" Then
                ''    'Ask for the Spouse's SIN number to create the record
                ''    'glbChgTermDate = ""
                ''    glbSpouseSIN = ""
                ''    frmMsgTerm.PenTermDate = "SpouseSIN"
                ''    frmMsgTerm.Show 1
                ''End If
            Else
                'Whenever there is a change of DB Pension beneficiary in IHR, the Pension Master Beneficiary should be updated too. We are removing the Please investigate message.
                ''loop all records for this type
                'Do While Not rsPen.EOF
                '    If rsPen("PE_BEN_NAME") = rsHRBenF("BD_BNAME") Then
                '        If IsDate(rsPen("PE_BEN_DOB")) And IsDate(rsHRBenF("BD_DOB")) Then
                '            If CVDate(rsPen("PE_BEN_DOB")) = CVDate(rsHRBenF("BD_DOB")) Then
                '                'Found it, no update
                '                Exit Sub
                '            End If
                '        End If
                '    End If
                '    rsPen.MoveNext
                'Loop
                'MsgBox "Spouse beneficiary is different between the Employee's benefits " & Chr(10) & "and Pension.  Please investigate."
                'Exit Sub
            End If
            If xUpdFlag = "NEW" Then
                rsPen.AddNew
                rsPen("PE_SIN") = xSIN
                rsPen("PE_PENSIONTYPE") = xPenType '"DBS"
                rsPen("PE_EMPNBR") = rsEmp("ED_EMPNBR")
                rsPen("PE_SURNAME") = rsEmp("ED_SURNAME")
                rsPen("PE_FNAME") = rsEmp("ED_FNAME")
                'Ticket #24275 Franks 08/27/2013
                rsPen("PE_AS_OF") = Date
                rsPen("PE_CURRENT") = -1
                xOrgCurrent = -1
            Else
                If IsNull(rsPen("PE_CURRENT")) Then
                    xOrgCurrent = 0
                Else
                    xOrgCurrent = rsPen("PE_CURRENT")
                End If
            End If
            'rsPen("PE_HRLYSAL") = xSalHly '"Salaried"
            'rsPen("PE_COUNTRY") = "CAN"
            rsPen("PE_SECTION") = xSection
            rsPen("PE_DIV") = rsEmp("ED_DIV")
            rsPen("PE_BEN_NAME") = rsHRBenF("BD_BNAME")
            rsPen("PE_BEN_RELATE") = rsHRBenF("BD_RELATE")
            If Not IsNull(rsHRBenF("BD_RELATE")) Then
                If UCase(rsHRBenF("BD_RELATE")) = "SPOUSE" Then
                    rsPen("PE_NOT_SPOUSE") = 0
                Else
                    rsPen("PE_NOT_SPOUSE") = 1
                End If
            End If
            rsPen("PE_BEN_DOB") = rsHRBenF("BD_DOB")
            rsPen("PE_DEATHDATE") = rsHRBenF("BD_DEATHDATE")
            rsPen("PE_LDATE") = Date '
            rsPen("PE_LTIME") = Time$
            rsPen("PE_LUSER") = glbUserID
            
            'Ticket #21021 Franks 02/28/2012
            rsPen("PE_SEP_AGREE") = rsHRBenF("BD_SEP_AGREE")
            rsPen("PE_SPOUSE_ENT") = rsHRBenF("BD_SPOUSE_ENT")
            rsPen("PE_SPE_DATE") = rsHRBenF("BD_SPE_DATE")
            
            'Ticket #21543 Franks 02/07/2012, Jerry & MZ asked to remove this logic
            'If Len(glbChgTermDate) > 0 Then
            If Len(glbSpouseSIN) > 0 Then
                rsPen("PE_SPOUSE_SIN") = glbSpouseSIN 'glbChgTermDate
            End If
            
            If Len(xLocSpoSIN) > 0 Then 'the previous logic
                rsPen("PE_SPOUSE_SIN") = xLocSpoSIN
            End If
            'Ticket #24275 Franks 08/27/2013 - begin
            rsPen("PE_PC") = rsHRBenF("BD_PC")
            
            '----current logic ------ start -------
            'o   If Date of Separation is entered and Spouse Entitled to Pension is checked, the Pension Beneficiary should be true.
            If Not IsNull(rsHRBenF("BD_SPE_DATE")) Then
                If rsHRBenF("BD_SPOUSE_ENT") Then
                    'If (IsNull(rsPen("PE_CURRENT")) Or rsPen("PE_CURRENT") = 0) Then
                    If xOrgCurrent = 0 Then
                        If IsNull(rsPen("PE_WAIVED_DATE")) Then
                            rsPen("PE_AS_OF") = Date
                            rsPen("PE_CURRENT") = -1
                        End If
                    End If
                Else
                'If Date of Separation is entered and Spouse Entitled to Pension is not checked
                    If Not xOrgCurrent = 0 Then
                        If IsNull(rsPen("PE_WAIVED_DATE")) Then
                            rsPen("PE_AS_OF") = Date
                            rsPen("PE_CURRENT") = 0
                        End If
                    End If
                End If
            End If
            'for 100% beneficiary, then change the PE_CURRENT back to True if it is False
            'If (IsNull(rsPen("PE_CURRENT")) Or rsPen("PE_CURRENT") = 0) And rsPen("PE_PC") = 1 Then
            If xOrgCurrent = 0 And rsPen("PE_PC") = 1 Then
                If IsNull(rsHRBenF("BD_DEATHDATE")) And IsNull(rsHRBenF("BD_SPE_DATE")) And IsNull(rsHRBenF("BD_END_DATE")) Then
                    If IsNull(rsPen("PE_WAIVED_DATE")) Then
                        rsPen("PE_AS_OF") = Date
                        rsPen("PE_CURRENT") = -1
                    End If
                End If
            End If
            
            'o   If End Date or Date of Death is entered, the Pension Beneficiary is false.
            If Not IsNull(rsHRBenF("BD_END_DATE")) Or Not IsNull(rsHRBenF("BD_DEATHDATE")) Then
                'If rsPen("PE_CURRENT") = 1 Then
                If Not xOrgCurrent = 0 Then
                        rsPen("PE_AS_OF") = Date
                        rsPen("PE_CURRENT") = 0
                End If
            End If
            If Not IsMissing(xBensChg) Then
                If xBensChg Then
                    If IsDate(xAsOfBens) Then rsPen("PE_AS_OF") = xAsOfBens
                End If
            End If
            '----current logic ------ end -------
            
            '''Ticket #24337 Franks 09/10/2013
            '''xBensChg = false, user may remove End Date, Date of Death and Date of Separation
            '''for 100% beneficiary, then change the PE_CURRENT back to True if it is False
            ''If (IsNull(rsPen("PE_CURRENT") = 0) Or rsPen("PE_CURRENT") = 0) And rsPen("PE_PC") = 1 Then
            ''    If IsNull(rsHRBenF("BD_DEATHDATE")) And IsNull(rsHRBenF("BD_SPE_DATE")) And IsNull(rsHRBenF("BD_END_DATE")) Then
            ''        If IsNull(rsPen("PE_WAIVED_DATE")) Then
            ''            rsPen("PE_AS_OF") = Date
            ''            rsPen("PE_CURRENT") = -1
            ''        End If
            ''    End If
            ''End If
            ''If Not IsMissing(xBensChg) Then
            ''    If xBensChg Then
            ''        rsPen("PE_AS_OF") = xAsOfBens
            ''        rsPen("PE_CURRENT") = xCurrentBens
            ''    Else
            ''        '''xBensChg = false, user may remove End Date, Date of Death and Date of Separation
            ''        '''for 100% beneficiary, then change the PE_CURRENT back to True if it is False
            ''        ''If rsPen("PE_CURRENT") = 0 And rsPen("PE_PC") = 1 Then
            ''        ''    If IsNull(rsHRBenF("BD_DEATHDATE")) And IsNull(rsHRBenF("BD_SPE_DATE")) And IsNull(rsHRBenF("BD_END_DATE")) Then
            ''        ''        rsPen("PE_AS_OF") = Date
            ''        ''        rsPen("PE_CURRENT") = -1
            ''        ''    End If
            ''        ''End If
            ''        'If rsPen("PE_CURRENT") = 0 Then
            ''            'If Not IsNull(rsHRBenF("BD_DEATHDATE")) Or Not IsNull(rsHRBenF("BD_SPE_DATE")) Or Not IsNull(rsHRBenF("BD_END_DATE")) Then
            ''            If Not IsNull(rsHRBenF("BD_END_DATE")) Or Not IsNull(rsHRBenF("BD_DEATHDATE")) Then
            ''                If Not rsPen("PE_CURRENT") = 0 Then
            ''                    rsPen("PE_AS_OF") = Date
            ''                End If
            ''                rsPen("PE_CURRENT") = 0
            ''            End If
            ''        'End If
            ''    End If
            ''End If
            '''Ticket #24275 Franks 08/27/2013 - end
            rsPen.Update
            rsPen.Close
            'add beneficiary master record - end
            
            'Pension Alert - Benficiary
            Call WFCPensionAlerts(rsEmp("ED_EMPNBR"), Date, "Check Beneficiary")
            
            rsHRBenF.MoveNext
        'End If
        Loop
    End If
    rsEmp.Close
End Sub


Public Sub WorkFlowUpdate(xEmpNo, xType, Optional xActDate)
Dim rsWF_Emp As New ADODB.Recordset
Dim rsWF_Master As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xSIN, xSurname, xFName, xPlant
    SQLQ = "SELECT ED_EMPNBR, ED_SIN, ED_SURNAME,ED_FNAME, ED_SECTION FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsEmp.EOF Then Exit Sub
    xSIN = rsEmp("ED_SIN")
    xSurname = rsEmp("ED_SURNAME")
    xFName = rsEmp("ED_FNAME")
    xPlant = rsEmp("ED_SECTION")
    rsEmp.Close
    
    SQLQ = "SELECT * FROM HRWORKFLOW_MASTER WHERE WK_WORKFLOW = '" & xType & "' ORDER BY WK_STEP "
    rsWF_Master.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do While Not rsWF_Master.EOF
        
        SQLQ = "SELECT * FROM HRWORKFLOW_EMPLOYEE WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND PE_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND PE_SECTION = '" & xPlant & "' "
        SQLQ = SQLQ & "AND PE_WORKFLOW = '" & xType & "' "
        SQLQ = SQLQ & "AND PE_STEP = " & rsWF_Master("WK_STEP") & " "
        rsWF_Emp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsWF_Emp.EOF Then
            rsWF_Emp.AddNew
            rsWF_Emp("PE_COMPNO") = "001"
            'rsWF_Emp("PE_COUNTRY") = "CAN"
            rsWF_Emp("PE_SIN") = xSIN
        End If
        rsWF_Emp("PE_EMPNBR") = xEmpNo
        rsWF_Emp("PE_SURNAME") = xSurname
        rsWF_Emp("PE_FNAME") = xFName
        rsWF_Emp("PE_SECTION") = xPlant
        rsWF_Emp("PE_WORKFLOW") = xType
        rsWF_Emp("PE_STEP") = rsWF_Master("WK_STEP")
        rsWF_Emp("PE_TASK") = rsWF_Master("WK_TASK")
        rsWF_Emp("PE_TARGET") = rsWF_Master("WK_TARGET")
        If Not IsMissing(xActDate) Then
            rsWF_Emp("PE_TARGET_DATE") = WFCPensionGetTargetDate(rsWF_Master("WK_TARGET"), xActDate)
        End If
        rsWF_Emp("PE_BYWHOM") = rsWF_Master("WK_BYWHOM")
        rsWF_Emp("PE_LDATE") = Date
        rsWF_Emp("PE_LTIME") = Time$
        rsWF_Emp("PE_LUSER") = glbUserID
        rsWF_Emp.Update
        rsWF_Emp.Close
        rsWF_Master.MoveNext
    Loop
    rsWF_Master.Close
    
End Sub
Public Function WFCPensionGetTargetDate(xDay, xActDate)
Dim I As Integer
Dim retVal
    retVal = xActDate
    If IsNumeric(xDay) Then
        If xDay > 0 Then
            I = 0
            'If start date is in weekend, find Friday as Start Date
            If Weekday(retVal) = 1 Or Weekday(retVal) = 7 Then
                retVal = DateAdd("D", -1, retVal)
                If Weekday(retVal) = 1 Or Weekday(retVal) = 7 Then
                    retVal = DateAdd("D", -1, retVal)
                End If
            End If
            Do While I < xDay
                If Weekday(retVal) = 1 Or Weekday(retVal) = 7 Then
                    '1 - Sunday, 7 - Saturday
                    'exclude weekend
                Else
                    I = I + 1
                End If
                retVal = DateAdd("D", 1, retVal)
            Loop
            'If end date is in weekend, find Monday as end Date
            If Weekday(retVal) = 1 Or Weekday(retVal) = 7 Then
                retVal = DateAdd("D", 1, retVal)
                If Weekday(retVal) = 1 Or Weekday(retVal) = 7 Then
                    retVal = DateAdd("D", 1, retVal)
                End If
            End If
        End If
    End If
    WFCPensionGetTargetDate = retVal
End Function
Public Sub WFCPensionAlerts(xEmpNo, xEventDate, xType, Optional xNewData, Optional xOldData, Optional xTERM_Seq, Optional IsDelete = "N")
Dim rsAlerts As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xSIN, xSurname, xFName, xPlant
    If IsMissing(xTERM_Seq) Then
        SQLQ = "SELECT ED_EMPNBR, ED_SIN, ED_SURNAME,ED_FNAME, ED_SECTION,ED_EMPTYPE FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_SIN, ED_SURNAME,ED_FNAME, ED_SECTION,ED_EMPTYPE FROM TERM_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_Seq & " "
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsEmp.EOF Then Exit Sub
        
    If IsNull(rsEmp("ED_EMPTYPE")) Then Exit Sub
    If Not rsEmp("ED_EMPTYPE") = "Y" Then
            Exit Sub
    End If
        
    xSIN = rsEmp("ED_SIN")
    xSurname = rsEmp("ED_SURNAME")
    xFName = rsEmp("ED_FNAME")
    xPlant = rsEmp("ED_SECTION")
    rsEmp.Close
    
    'Ticket #22009 Franks 05/09/2012 - begin
    If IsDelete = "ALL" Then
        SQLQ = "DELETE FROM HRP_PENSION_ALERTS WHERE PE_SIN = '" & xSIN & "' "
        gdbAdoIhr001.Execute SQLQ
        Exit Sub
    End If
    If IsDelete = "Y" Then
        SQLQ = "DELETE FROM HRP_PENSION_ALERTS WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND NOT PE_EVENT_TYPE = '" & xType & "' "
        gdbAdoIhr001.Execute SQLQ
        Exit Sub
    End If
    'Ticket #22009 Franks 05/09/2012 - end
    
    SQLQ = "SELECT * FROM HRP_PENSION_ALERTS WHERE PE_SIN = '" & xSIN & "' "
    SQLQ = SQLQ & "AND PE_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND PE_SECTION = '" & xPlant & "' "
    SQLQ = SQLQ & "AND PE_EVENT_DATE = " & Date_SQL(xEventDate) & " "
    SQLQ = SQLQ & "AND PE_EVENT_TYPE = '" & xType & "' "
    rsAlerts.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsAlerts.EOF Then
        rsAlerts.AddNew
        rsAlerts("PE_COMPNO") = "001"
        rsAlerts("PE_COUNTRY") = "CAN"
        rsAlerts("PE_SIN") = xSIN
    End If
    rsAlerts("PE_EMPNBR") = xEmpNo
    rsAlerts("PE_SURNAME") = xSurname
    rsAlerts("PE_FNAME") = xFName
    rsAlerts("PE_SECTION") = xPlant
    rsAlerts("PE_EVENT_DATE") = xEventDate
    rsAlerts("PE_EVENT_TYPE") = xType
    If Not IsMissing(xOldData) Then
        rsAlerts("PE_OLD_VALUE") = xOldData
    End If
    If Not IsMissing(xNewData) Then
        rsAlerts("PE_NEW_VALUE") = xNewData
    End If
    rsAlerts("PE_LDATE") = Date
    rsAlerts("PE_LTIME") = Time$
    rsAlerts("PE_LUSER") = glbUserID
    rsAlerts.Update
    rsAlerts.Close
    
End Sub

Public Sub WFCOtherPenUpt(xEmpNo, xSIN, xYear, xSalHly, xPenMainType, xDBStatus, xEffDate, Optional xLastDay, Optional xDBDC)
Dim rsEmp As New ADODB.Recordset
Dim rsEmOther2 As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim rsHRBeFicia As New ADODB.Recordset
Dim SQLQ As String
Dim xExcludeList As String
Dim xPenType As String
    If Len(xPenMainType) > 0 Then
        xExcludeList = "'" & xPenMainType & "'"
    Else
        xExcludeList = ""
    End If
    Do While True
        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND PE_YEAR_DATE <= " & xYear & " "
        If Len(xSalHly) > 0 Then
            SQLQ = SQLQ & "AND PE_HRLYSAL = '" & xSalHly & "' "
        End If
        If Len(xExcludeList) > 0 Then
            SQLQ = SQLQ & "AND NOT PE_PENSIONTYPE IN (" & xExcludeList & ") "
        End If
        If IsMissing(xDBDC) Then
            SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
        Else
            SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = '" & xDBDC & "' "
        End If
        SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
        If rsPen.State <> 0 Then rsPen.Close
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsPen.EOF Then
            Exit Sub
        End If
        xPenType = rsPen("PE_PENSIONTYPE")
        xSalHly = rsPen("PE_HRLYSAL")
        If Len(xExcludeList) > 0 Then
            xExcludeList = xExcludeList & ",'" & xPenType & "'"
        Else
            xExcludeList = "'" & xPenType & "'"
        End If
        
        'Reopen for this
        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' "
        If Len(xSalHly) > 0 Then
            SQLQ = SQLQ & "AND PE_HRLYSAL = '" & xSalHly & "' "
        End If
        SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
        If rsPen.State <> 0 Then rsPen.Close
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'No Pension Master record found, add it first
        If rsPen.EOF Then
            'Pen Master Audit -  Ticket #19954 - begin
            UpdPenAudit = True
            UpdPenAudDirect = False
            Call PenMasterAuditOldValSetup("Blank")
            toTYPE = "Add"
            'Pen Master Audit -  Ticket #19954 - end
            rsPen.Close
            Call WFCPensionMaster(xEmpNo, "N", xDBStatus, , xYear, , "Y", xPenType, xSalHly)
            'reopen this recordset again
            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Else
            'Pen Master Audit -  Ticket #19954 - begin
            Call PenMasterAuditOldValSetup("CurValues", rsPen)
            toTYPE = "Change"
            'Pen Master Audit -  Ticket #19954 - end
        End If

        rsPen("PE_DB_STATUS") = xDBStatus '"R"
        rsPen("PE_DB_STATUS_DATE") = xEffDate
        If Not IsMissing(xLastDay) Then
            If IsNull(rsPen("PE_EXIT_DATE")) Then
                If IsDate(xLastDay) Then '
                    rsPen("PE_EXIT_DATE") = xLastDay
                End If
            End If
        End If
        rsPen("PE_LUSER") = glbUserID
        rsPen("PE_LTIME") = Time$
        rsPen("PE_LDATE") = Date
        rsPen.Update
        'Pen Master Audit Ticket #19954
        Call AUDIT_PENSION_MASTER(rsPen)
        rsPen.Close
    Loop
    
End Sub
Public Sub WFCPensionMasUpt(xEmpNo, xUptType, xUptVal, Optional xUptOldVal, Optional xActionYear, Optional xPAData, Optional xOldUnion = "")
'see Pension Phase II.doc in W:\2008 Projects\Pension\Pension Phase II folder
Dim rsEmp As New ADODB.Recordset
Dim rsEmOther2 As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim rsHRBeFicia As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xSection, xPenType, xDBStatus, xSalHly, xDIV As String, xDivNew As String
Dim xDOB, xEarlyRet, xNorRet, xLateRet, xDATE, xYear, xSIN, xUnion
Dim xPenExitDate, xOldValue, xNewValue
Dim xTmp1, xTmp2
Dim xTermCode As String, xTermDate, xRetireDate
Dim xCurNOCC
Dim xlTmpDateFr, xlTmpDateTo, xContServ, xCreditService, xEarnedPen
Dim xDOH, xLastDay, xCurStatus, xFirstPenDay, xlCreditServ, xBenRate, xPAAmt
Dim xFDate, xtmpdate
Dim xUpdFlag As Boolean
Dim AllUpt As Boolean
Dim xNewRec As Boolean 'Ticket #19954
Dim xSERYN As String, xSupRate '''Ticket #20717
Dim xComeInStatus

    xPenExitDate = ""
    If xUptType = "PenExitDate" Then
        If IsDate(xUptVal) Then
            xPenExitDate = xUptVal
        Else
            Exit Sub
        End If
    End If
    
    UpdPenAudDirect = False 'Ticket #19954
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        If IsNull(rsEmp("ED_EMPTYPE")) Then Exit Sub
        If Not rsEmp("ED_EMPTYPE") = "Y" Then
            If Not IsDate(xPenExitDate) Then
                Exit Sub
            End If
        End If
        If IsNull(rsEmp("ED_SIN")) Then Exit Sub
        If IsNull(rsEmp("ED_SECTION")) Then Exit Sub
        If IsNull(rsEmp("ED_DOH")) Then Exit Sub
        If IsNull(rsEmp("ED_DOB")) Then Exit Sub
        If IsNull(rsEmp("ED_ORG")) Then Exit Sub
        If IsNull(rsEmp("ED_ELIGIBLE")) Then Exit Sub
        xDIV = rsEmp("ED_DIV")
        
        xSIN = rsEmp("ED_SIN")
        xSection = rsEmp("ED_SECTION")
        xUnion = rsEmp("ED_ORG")
        xDOH = rsEmp("ED_DOH")
        If IsMissing(xActionYear) Then
            xYear = Year(rsEmp("ED_ELIGIBLE"))
        Else
            xYear = xActionYear
        End If
        
        '''Ticket #20717
        xSERYN = ""
        If Not IsNull(rsEmp("ED_HIRECODE")) Then
            xSERYN = rsEmp("ED_HIRECODE")
        End If
        
        xPenType = getDBType(xSection, xUnion, "PenType", rsEmp("ED_DOH")) 'Ticket #26707 Franks 02/25/2015
        xSalHly = getDBType(xSection, xUnion, "HlySal")
        If Len(xPenType) = 0 Then
            Exit Sub
        End If
        If Len(xSalHly) = 0 Then
            Exit Sub
        End If
        
        xDBStatus = "H"
        
        If Not IsNull(rsEmp("ED_DOB")) Then
            xDOB = rsEmp("ED_DOB")
        Else
            xDOB = Date
        End If
        
        If xUptType = "PenExitDate" Then
            'update Pension Exit Date - begin
            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            
            'No Pension Master record found, add it first
            If rsPen.EOF Then
                'Pen Master Audit -  Ticket #19954 - begin
                UpdPenAudit = True
                Call PenMasterAuditOldValSetup("Blank")
                toTYPE = "Add"
                'Pen Master Audit -  Ticket #19954 - end
                rsPen.Close
                Call WFCPensionMaster(xEmpNo, "Y", , , xYear, "PenExitDate")
                'reopen this recordset again
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Else
                'Pen Master Audit -  Ticket #19954 - begin
                Call PenMasterAuditOldValSetup("CurValues", rsPen)
                toTYPE = "Change"
                'Pen Master Audit -  Ticket #19954 - end
            End If
            If Not rsPen.EOF Then
                rsPen("PE_EXIT_DATE") = xPenExitDate
                rsPen("PE_DB_STATUS_DATE") = xPenExitDate
                'rsPen("PE_DB_STATUS") = "N"
                'Ticket #23361 Franks 03/11/2011
                rsPen("PE_DB_STATUS") = "T"
                'Pen Master Audit Ticket #19954
                Call AUDIT_PENSION_MASTER(rsPen)
                rsPen.Update
            End If
            rsPen.Close
            'update Pension Exit Date  - end

            'Pension Alert - Benficiary
            Call WFCPensionAlerts(xEmpNo, Date, "Pension Eligibility Changed", xPenExitDate)
        End If
        
        If xUptType = "Re-Active from a Leave" Then
        'If the LOA return year is the same year as the Pension Master, change the status to
        'Active with the new Effective Date equal to the Effective As Of date. If the LOA year is
        'greater than the year of the suspended record, create a new Pension Master.
        'Hourly employee only
            'Ticket #19954 Franks 03/28/2011
            ' Salaried & Hourly: Return from a leave changes the pension status to "A" and the Status Effective Date equals the "New Employment Effective Date as of".
            'If xSalHly = "Hourly" Then
            
                'Ticket #21597 Franks 05/01/2012
                xComeInStatus = ""
                If Not IsMissing(xUptOldVal) Then
                    xComeInStatus = xUptOldVal
                End If
            
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                'No Pension Master record found, add it first
                If rsPen.EOF Then
                    'Pen Master Audit -  Ticket #19954 - begin
                    UpdPenAudit = True
                    Call PenMasterAuditOldValSetup("Blank")
                    toTYPE = "Add"
                    'Pen Master Audit -  Ticket #19954 - end
                    rsPen.Close
                    Call WFCPensionMaster(xEmpNo, "Y", , , xYear)
                    'reopen this recordset again
                    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                    SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                    SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                    If rsPen.State <> 0 Then rsPen.Close
                    rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Else
                    'Pen Master Audit -  Ticket #19954 - begin
                    Call PenMasterAuditOldValSetup("CurValues", rsPen)
                    toTYPE = "Change"
                    'Pen Master Audit -  Ticket #19954 - end
                End If
                If Not rsPen.EOF Then
                    xFDate = xUptVal
                    If Len(xComeInStatus) > 0 Then
                        If Not (xComeInStatus = rsPen("PE_DB_STATUS")) Then 'Ticket #21597 Franks 05/01/2012
                            rsPen("PE_DB_STATUS") = Left(xComeInStatus, 1)
                            If IsDate(xFDate) Then rsPen("PE_DB_STATUS_DATE") = xFDate
                        Else
                            'nothing change
                        End If
                    Else
                        rsPen("PE_DB_STATUS") = "A"
                        If IsDate(xFDate) Then rsPen("PE_DB_STATUS_DATE") = xFDate
                    End If
                    
                    'Ticket #21788 Franks 03/26/2012
                    '"   Remove the Benefit Rate
                    If xSalHly = "Hourly" Then
                        rsPen("PE_BENEFIT_RATE") = Null
                    End If
                    
                    rsPen.Update
                    'Pen Master Audit Ticket #19954
                    Call AUDIT_PENSION_MASTER(rsPen)
                End If
                rsPen.Close
            'End If
            
        End If
        
        If xUptType = "Temporary Layoff" Then
            
            'Ticket #21597 Franks 05/01/2012
            xComeInStatus = ""
            If Not IsMissing(xUptOldVal) Then
                xComeInStatus = xUptOldVal
            End If
            
            'Ticket #21822 Franks 03/30/2012, this function needs to work for both Hourly and Salaried
            'it worked for Hourly only before
            'If xSalHly = "Hourly" Then
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                'No Pension Master record found, add it first
                If rsPen.EOF Then
                    rsPen.Close
                    Call WFCPensionMaster(xEmpNo, "Y", , , xYear)
                    'reopen this recordset again
                    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                    SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                    SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                    If rsPen.State <> 0 Then rsPen.Close
                    rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                End If
                If Not rsPen.EOF Then
                    xFDate = xUptVal
                    xLastDay = xFDate
                    If Len(xComeInStatus) > 0 Then
                        If Not (xComeInStatus = rsPen("PE_DB_STATUS")) Then 'Ticket #21597 Franks 05/01/2012
                            rsPen("PE_DB_STATUS") = Left(xComeInStatus, 1)
                            If IsDate(xFDate) Then rsPen("PE_DB_STATUS_DATE") = xFDate
                        Else
                            'nothing change
                        End If
                    Else
                        rsPen("PE_DB_STATUS") = "S"
                        If IsDate(xFDate) Then rsPen("PE_DB_STATUS_DATE") = xFDate
                    End If
                    
                    If xSalHly = "Hourly" Then
                        'Ticket #21788 Franks 03/26/2012 - begin
                        If Not IsNull(rsPen("PE_NOGC")) Then
                            xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xLastDay, rsPen("PE_NOGC"))
                        Else
                            xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xLastDay)
                        End If
                        If xBenRate > 0 Then
                            rsPen("PE_BENEFIT_RATE") = xBenRate
                        End If
                        'Ticket #21788 Franks 03/26/2012 - end
                    End If
                    
                    rsPen.Update
                End If
                rsPen.Close
            'End If
        End If
        
        If xUptType = "DOH_Change" Then
            'If the Pension Eligibility Flag is "Y" and there is only one DB Pension Master record for
            'that employee, change the DB Entry Date to equal the Original Date of Hire. Otherwise,
            'create a Pension Alert with the Type of Event = Original Date of Hire Changed.
            'Also update the Old Value and New Value.
            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsPen.EOF Then
                If rsPen.RecordCount = 1 Then
                    If IsDate(xUptVal) Then
                        'Pen Master Audit -  Ticket #19954 - begin
                        Call PenMasterAuditOldValSetup("CurValues", rsPen)
                        toTYPE = "Change"
                        'Pen Master Audit -  Ticket #19954 - end
                        rsPen("PE_PEN_ENTRY_DATE") = xUptVal
                        'rsPen("PE_DB_STATUS_DATE") = xUptVal
                        rsPen.Update
                        'Pen Master Audit Ticket #19954
                        Call AUDIT_PENSION_MASTER(rsPen)
                        'update Pension Eligible date as well
                        rsEmp("ED_ELIGIBLE") = xUptVal
                        rsEmp.Update
                    End If
                End If
            Else
                'Pension Alert
                Call WFCPensionAlerts(xEmpNo, Date, "Original Date of Hire Changed", xUptVal, xUptOldVal)
            End If
            rsPen.Close
        End If
        
        'Ticket #18566
        If xUptType = "Retirement" Then
        'Pension Master:  New record with Year of retirement date, Pension Entry Date and the other fields
        'above the first line remain the same as the previous record . If there is no DB Pension Master record
        'for that year, create one. - Retirement Process.docx
            If xSalHly = "Hourly" Then
                xCurNOCC = GetCurNOCC(xEmpNo)
            End If
            xLastDay = xUptOldVal
            xRetireDate = xUptVal 'xTERMDATE = xUptVal
            
            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            AllUpt = True
            
            'No Pension Master record found, add it first
            If rsPen.EOF Then
                'Pen Master Audit -  Ticket #19954 - begin
                UpdPenAudit = True
                Call PenMasterAuditOldValSetup("Blank")
                toTYPE = "Add"
                'Pen Master Audit -  Ticket #19954 - end
                rsPen.Close
                Call WFCPensionMaster(xEmpNo, "N", , , xYear, , "Y")
                'reopen this recordset again
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            Else
                'Pen Master Audit -  Ticket #19954 - begin
                Call PenMasterAuditOldValSetup("CurValues", rsPen)
                toTYPE = "Change"
                'Pen Master Audit -  Ticket #19954 - end
            End If
            
            xDBStatus = "R"
            rsPen("PE_DB_STATUS") = xDBStatus
            
            'Credit Service, Cont. Service
            'Max is 12 months, if new hire in this year, convert the days to months
            xlTmpDateFr = CVDate("Jan 1, " & xYear)
            xlTmpDateTo = xLastDay 'xTERMDATE
    
            If IsDate(rsEmp("ED_DOH")) Then
                If CVDate(rsEmp("ED_DOH")) > CVDate(xlTmpDateFr) Then
                    xlTmpDateFr = rsEmp("ED_DOH")
                End If
            End If
            xContServ = DateDiff("D", xlTmpDateFr, xlTmpDateTo) + 1 'xLastDay)
            If xContServ > 365 Then
                xContServ = 365
            End If
            If xContServ < 0 Then xContServ = 0
            xContServ = Round((xContServ / (365 / 12)), 2)
            If xContServ >= 11.96 And xContServ <= 12 Then
                xContServ = 12
            End If
            If xContServ <= 0.07 Then xContServ = 0
            xContServ = Round(xContServ, 1) 'Ticket #23704 Franks 05/06/2013
            
            xCreditService = xContServ
            
            If xSalHly = "Salaried" Then
                xCreditService = xContServ
            Else
                'Hourly employees
                'Check Status change history
                'xLastDay = xTERMDATE
                xFirstPenDay = xlTmpDateFr
                If IsDate(rsEmp("ED_ELIGIBLE")) Then
                    'xPenStartDate = rsEmp("ED_ELIGIBLE")
                    If Not (CVDate(rsEmp("ED_ELIGIBLE")) <= CVDate(xFirstPenDay)) Then 'xFirstDay
                        xFirstPenDay = rsEmp("ED_ELIGIBLE")
                    End If
                End If
                xCurStatus = rsEmp("ED_EMP")
            'If xEmpNo = 12352191 Then ' 12109104 Then
            'Debug.Print ""
            'End If
                Call AddEmpStatus(xEmpNo, xYear, xLastDay, xCurStatus, xFirstPenDay)
                'Calculate the Credited Service - Begin
                xlCreditServ = CalcCreditService(xEmpNo, xYear, xUnion, xSection, xDOH, xLastDay)
                'If xlCreditServ < xCreditService Then
                If xlCreditServ <= 12 Then
                    xCreditService = xlCreditServ
                End If
                'Calculate the Credited Service - End

            End If
            
            xEarnedPen = 0
            'If the retirement creates a Pension Master and there is no earned pension calculated,
            'update the field with zero. Otherwise, they need to enter the zero anytime they go into that record.
            If IsNull(rsPen("PE_YEAR_AMOUNT")) Then
                rsPen("PE_YEAR_AMOUNT") = 0
            End If
            If xSalHly = "Salaried" Then
                If Not IsMissing(xPAData) Then 'xPAData -  Pensionable Earnings
                    xEarnedPen = xPAData * 0.009
                    rsPen("PE_YEAR_AMOUNT") = xEarnedPen
                End If
            End If
            If xSalHly = "Hourly" Then
                If Not IsNull(rsPen("PE_NOGC")) Then
                    xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xLastDay, rsPen("PE_NOGC"))
                Else
                    xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xLastDay)
                End If

                If xBenRate > 0 Then
                    If IsNull(rsPen("PE_BENEFIT_RATE")) Then
                        rsPen("PE_BENEFIT_RATE") = xBenRate
                    Else
                        If rsPen("PE_BENEFIT_RATE") = 0 Then
                            rsPen("PE_BENEFIT_RATE") = xBenRate
                        End If
                    End If
                    rsPen("PE_YEAR_AMOUNT") = xCreditService * xBenRate
                    xEarnedPen = rsPen("PE_YEAR_AMOUNT")
                End If

                '''Ticket #20717
                If xSERYN = "Y" Then
                    If Not IsNull(rsPen("PE_NOGC")) Then
                        xSupRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xLastDay, rsPen("PE_NOGC"), "Y")
                    Else
                        xSupRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xLastDay, , "Y")
                    End If
    
                    If xSupRate > 0 Then
                        If IsNull(rsPen("PE_SUPPLE_RATE")) Then
                            rsPen("PE_SUPPLE_RATE") = xSupRate
                        Else
                            If rsPen("PE_SUPPLE_RATE") = 0 Then
                                rsPen("PE_SUPPLE_RATE") = xSupRate
                            End If
                        End If
                    End If
                End If
            End If
            rsPen("PE_CONT_SERV") = xContServ
            rsPen("PE_CREDITED_SERV") = xCreditService
            rsPen("PE_DB_STATUS_DATE") = xRetireDate  'xTERMDATE
            If IsNull(rsPen("PE_EXIT_DATE")) Then
                If IsDate(xLastDay) Then '
                    rsPen("PE_EXIT_DATE") = xLastDay
                End If
            End If
            rsPen("PE_LUSER") = glbUserID
            rsPen("PE_LTIME") = Time$
            rsPen("PE_LDATE") = Date
            rsPen.Update
            'Pen Master Audit Ticket #19954
            Call AUDIT_PENSION_MASTER(rsPen)
            rsPen.Close
            
            'Create PA Master and PA Detail on Employee Retirement - Begin
            'If xEarnedPen > 0 Then 'Not IsMissing(xPAData) Then
            If Not IsMissing(xPAData) Then
                Call WFCPAMaster(xEmpNo, xSIN, xYear, xSalHly, "Y", xPAData) ' xEarnedPen)
            End If
            'Create PA Master and PA Detail on Employee Retirement - End
            
            'Retire Other DB Pensions
            'One employee can have one DBS plus other DB pensions, such as DBKIPL
            'Employee Dan Dubblestyne had DBS and DBKIPL pensions
            'move this function into Retirement Process in frmERetirement
            'Call WFCOtherPenUpt(xEmpNo, xSIN, xYear, xSalHly, xPenType, "R", xRetireDate, xLastDay)
        End If
        'Retirement - end
        
        If xUptType = "Termination" Then
            If xSalHly = "Hourly" Then
                xCurNOCC = GetCurNOCC(xEmpNo)
            End If
            
            'Ticket #25054 Franks 02/12/2014 - begin
            '. If the most recent pension status is "C" or "D", the termination process shouldn't update the Pension Master.
            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_YEAR_DATE <= " & xYear & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsPen.EOF Then
                If rsPen("PE_DB_STATUS") = "C" Or rsPen("PE_DB_STATUS") = "D" Then
                    Exit Sub
                End If
            End If
            rsPen.Close
            'Ticket #25054 Franks 02/12/2014 - end
            
            xTermCode = xUptOldVal
            xTermDate = xUptVal
            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
            If rsPen.State <> 0 Then rsPen.Close
            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            AllUpt = True
            
            'Ticket #22591 Franks 10/12/2012
            '"   Terminated employees who were hired in the same year should not delete the Pension Master files. Their status should be "T" and their pension calculated just like all other employees.
            ''''On Termination, if the Pension Status is New Hire and the year in the
            ''''Membership Entry Date equals the year of the DOT, delete the pension master
            '''If Not rsPen.EOF Then 'find the record
            '''    If rsPen("PE_DB_STATUS") = "H" Then 'New hire
            '''        If IsDate(rsPen("PE_PEN_ENTRY_DATE")) Then
            '''            If IsDate(xTERMDATE) Then
            '''                If Year(xTERMDATE) = Year(rsPen("PE_PEN_ENTRY_DATE")) Then
            '''                    'If Year(xTermDate) = Year(rsEmp("ED_DOH")) Then
            '''                    If CVDate(xTERMDATE) > CVDate(rsEmp("ED_DOH")) Then
            '''                        rsPen.Delete
            '''                        rsPen.Close
            '''                        Exit Sub
            '''                    End If
            '''                End If
            '''            End If
            '''        End If
            '''    End If
            '''End If
            
            'No Pension Master record found, add it first
            xNewRec = False
            If rsPen.EOF Then
                'Pen Master Audit -  Ticket #19954 - begin
                UpdPenAudit = True
                Call PenMasterAuditOldValSetup("Blank")
                toTYPE = "Add"
                'Pen Master Audit -  Ticket #19954 - end
                rsPen.Close
                Call WFCPensionMaster(xEmpNo, "Y", , , xYear)
                'reopen this recordset again
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                xNewRec = True
            Else
                'Pen Master Audit -  Ticket #19954 - begin
                Call PenMasterAuditOldValSetup("CurValues", rsPen)
                toTYPE = "Change"
                'Pen Master Audit -  Ticket #19954 - end
            End If
            
            'Termination - If the DOH => 2 years from the DOT, update Pension Master with  Status = T,
            'calculate continuous and credited service and update the Benefit Rate and NOGC (for hourly).
            'If they have less than 2 years, the status should be N. For both statuses update the status effective
            'date  and membership exit date to dot. If termination Reason = RET, the status code is "R".
            'Rest applies as above.
            ''Ticket #22285 Franks 07/16/2012 - do not change the Pension Status to "N". The Pension Status should be "T" instead.
            ''If (DateDiff("d", rsEmp("ED_DOH"), CVDate(xTERMDATE)) / 365) >= 2 Then
            ''    xDBStatus = "T"
            ''Else
            ''    xDBStatus = "N"
            ''End If
            xDBStatus = "T"
            
            If xTermCode = "RET" Then
                xDBStatus = "R"
            End If
            If xTermCode = "DECD" Then
                xDBStatus = "D"
            End If
            
            '------------Terminated Pension Update - begin
            If rsPen("PE_DB_STATUS") = "C" Or rsPen("PE_DB_STATUS") = "R" Or rsPen("PE_DB_STATUS") = "N" Or rsPen("PE_DB_STATUS") = "T" Then
                ' if the status code equals C or R. For these records, only update the Cont. Service, Credited Service and Earned Pension
                AllUpt = False
            Else
                rsPen("PE_DB_STATUS") = xDBStatus
            End If
            
            'Credit Service, Cont. Service
            'Max is 12 months, if new hire in this year, convert the days to months
            xlTmpDateFr = CVDate("Jan 1, " & xYear)
            xlTmpDateTo = xTermDate
    
            If IsDate(rsEmp("ED_DOH")) Then
                If CVDate(rsEmp("ED_DOH")) > CVDate(xlTmpDateFr) Then
                    xlTmpDateFr = rsEmp("ED_DOH")
                End If
            End If
            xContServ = DateDiff("D", xlTmpDateFr, xlTmpDateTo) + 1 'xLastDay)
            If xContServ > 365 Then
                xContServ = 365
            End If
            If xContServ < 0 Then xContServ = 0
            xContServ = Round((xContServ / (365 / 12)), 2)
            If xContServ >= 11.96 And xContServ <= 12 Then
                xContServ = 12
            End If
            If xContServ <= 0.07 Then xContServ = 0
            xContServ = Round(xContServ, 1) 'Ticket #23704 Franks 05/06/2013
            
            xCreditService = xContServ
            
            If xSalHly = "Salaried" Then
                xCreditService = xContServ
            Else
                'Hourly employees
                'Check Status change history
                xLastDay = xTermDate
                xFirstPenDay = xlTmpDateFr
                If IsDate(rsEmp("ED_ELIGIBLE")) Then
                    'xPenStartDate = rsEmp("ED_ELIGIBLE")
                    If Not (CVDate(rsEmp("ED_ELIGIBLE")) <= CVDate(xFirstPenDay)) Then 'xFirstDay
                        xFirstPenDay = rsEmp("ED_ELIGIBLE")
                    End If
                End If
                xCurStatus = rsEmp("ED_EMP")
            'If xEmpNo = 12352191 Then ' 12109104 Then
            'Debug.Print ""
            'End If
                Call AddEmpStatus(xEmpNo, xYear, xLastDay, xCurStatus, xFirstPenDay)
                'Calculate the Credited Service - Begin
                xlCreditServ = CalcCreditService(xEmpNo, xYear, xUnion, xSection, xDOH, xLastDay)
                'If xlCreditServ < xCreditService Then
                If xlCreditServ <= 12 Then
                    xCreditService = xlCreditServ
                End If
                'Calculate the Credited Service - End

            End If
            
            xEarnedPen = 0
            If xSalHly = "Hourly" Then
                If Not IsNull(rsPen("PE_NOGC")) Then
                    xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate, rsPen("PE_NOGC"))
                Else
                    xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate)
                End If

                If xBenRate > 0 Then
                    If IsNull(rsPen("PE_BENEFIT_RATE")) Then
                        rsPen("PE_BENEFIT_RATE") = xBenRate
                    Else
                        If rsPen("PE_BENEFIT_RATE") = 0 Then
                            rsPen("PE_BENEFIT_RATE") = xBenRate
                        End If
                    End If
                    rsPen("PE_YEAR_AMOUNT") = xCreditService * xBenRate
                    xEarnedPen = rsPen("PE_YEAR_AMOUNT")
                End If
                'Ticket #21012 Franks 09/29/2011
                'When terminating an employee who has Pension Eligibility = “Yes”,
                'update the Benefit Rate using the Pension Rate Master and the check the Freeze Rate.
                rsPen("PE_FREEZE_RATE") = 1
                
                '''Ticket #20717
                If xSERYN = "Y" Then
                    If Not IsNull(rsPen("PE_NOGC")) Then
                        xSupRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate, rsPen("PE_NOGC"), "Y")
                    Else
                        xSupRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate, , "Y")
                    End If
    
                    If xSupRate > 0 Then
                        If IsNull(rsPen("PE_SUPPLE_RATE")) Then
                            rsPen("PE_SUPPLE_RATE") = xSupRate
                        Else
                            If rsPen("PE_SUPPLE_RATE") = 0 Then
                                rsPen("PE_SUPPLE_RATE") = xSupRate
                            End If
                        End If
                    End If
                End If

            End If
            If Not IsMissing(xPAData) Then
                'Earned Pen. for Salaried
                If xSalHly = "Salaried" Then
                    xPAAmt = 0
                    If Len(xPAData) > 3 Then
                        xPAAmt = Mid(xPAData, 4, Len(xPAData) - 3)
                    End If
                    rsPen("PE_YEAR_AMOUNT") = xPAAmt * 0.009
                End If
            End If
            
            rsPen("PE_CONT_SERV") = xContServ
            rsPen("PE_CREDITED_SERV") = xCreditService
            
            If AllUpt Then 'none exited terminated record
                If IsDate(xTermDate) Then
                    If xTermCode = "RET" Then
                        'If Term Reason is RET, the Status Effective Date in the Pension Master
                        'would equal the first day of the month following. ONLY for this reason code
                        xtmpdate = DateAdd("M", 1, xTermDate)
                        xtmpdate = CVDate(MonthName(month(xtmpdate)) & " 1," & Str(Year(xtmpdate)))
                        rsPen("PE_DB_STATUS_DATE") = xtmpdate
                    Else
                        rsPen("PE_DB_STATUS_DATE") = xTermDate
                    End If
                    rsPen("PE_EXIT_DATE") = xTermDate
                End If
                If Len(xCurNOCC) > 0 Then
                    rsPen("PE_NOGC") = xCurNOCC
                End If
            End If
            
            'Ticket #19954 Franks 03/22/2011 - begin
            'Transaction processing needs to be changed. During termination if the only other Pension Master record
            'has a status "X" and a Calculated Pension entered, the Term Deferred Pension Master needs to have
            'the Calculated Pension from the "X" status to the Term Deferred status. (The "X" status's Calculated Pension should be zero'd out.)
            ''check the previous pension record if there is "X" status
            If xNewRec Then 'new recrod only
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE < " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                'SQLQ = SQLQ & "AND PE_DB_STATUS = 'X' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsTemp.State <> 0 Then rsTemp.Close
                rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsTemp.EOF Then
                    If Not IsNull(rsTemp("PE_DB_STATUS")) Then
                        If rsTemp("PE_DB_STATUS") = "X" Then
                            If Not IsNull(rsTemp("PE_ANNDEFERRED")) Then
                                If rsTemp("PE_ANNDEFERRED") > 0 Then
                                    rsPen("PE_ANNDEFERRED") = rsTemp("PE_ANNDEFERRED")
                                    rsTemp("PE_ANNDEFERRED") = 0
                                    rsTemp.Update
                                End If
                            End If
                        End If
                    End If
                End If
                rsTemp.Close
            End If
            'Ticket #19954 Franks 03/22/2011 - end

            If IsMissing(xUserID) Then
                rsPen("PE_LUSER") = glbUserID
            Else
                rsPen("PE_LUSER") = xUserID
            End If
            rsPen("PE_LTIME") = Time$
            rsPen("PE_LDATE") = Date
            rsPen.Update
            'Pen Master Audit Ticket #19954
            Call AUDIT_PENSION_MASTER(rsPen)
            rsPen.Close
            'update HR Status dates - Begin
            If IsDate(rsEmp("ED_DOB")) Then
                'Early Retirement Date
                If Not IsDate(rsEmp("ED_EARLYR")) Then
                    rsEmp("ED_EARLYR") = WFCPenEmpRetireDate(55, rsEmp("ED_DOB"))
                End If
                'Normal  Retirement Date
                If Not IsDate(rsEmp("ED_NORMALR")) Then
                    rsEmp("ED_NORMALR") = WFCPenEmpRetireDate(65, rsEmp("ED_DOB"))
                End If
                'Early Retirement Date
                If Not IsDate(rsEmp("ED_LATESTR")) Then
                    rsEmp("ED_LATESTR") = WFCPenEmpRetireDate(71, rsEmp("ED_DOB"))
                End If
                rsEmp.Update
                'xUptVal Pension Termnation date
                If IsDate(xTermDate) Then
                    SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo
                    If rsEmOther2.State <> 0 Then rsEmOther2.Close
                    rsEmOther2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsEmOther2.EOF Then
                        rsEmOther2.AddNew
                        rsEmOther2("ER_EMPNBR") = xEmpNo
                    End If
                    rsEmOther2("ER_PENSIONDATE4") = xTermDate 'Date of Termination
                    If xTermCode = "DECD" Then
                        rsEmOther2("ER_PENSIONDATE5") = xTermDate 'Date of Death
                    End If
                    rsEmOther2.Update
                End If
            End If
            
            'update HR Status dates - End
            '------------Terminated Pension Update - end
                
            'Create PA Master and PA Detail on Employee Termination - Begin
            If Not IsMissing(xPAData) Then
                'for Salaried xPAData
                If xSalHly = "Salaried" Then
                    xPAAmt = 0
                    If Len(xPAData) > 3 Then
                        xPAAmt = Mid(xPAData, 4, Len(xPAData) - 3)
                    End If
                    Call WFCPAMaster(xEmpNo, xSIN, xYear, xSalHly, "Y", xPAAmt)
                End If
                'for Hourly
                If xSalHly = "Hourly" Then
                    Call WFCPAMaster(xEmpNo, xSIN, xYear, xSalHly, "Y", xEarnedPen)
                End If
            End If
            'Create PA Master and PA Detail on Employee Termination - End
        End If
        '--------------Termination - end
        
        If xUptType = "Transfer Out" Then
        'Hourly Employees:
        '   Update the Pension Master with the Exit Date for the DB Pension.
        'Salaried Employees:
        '   If the Transfer To Division (use the Division Master to determine this) is still within the same Country, no Pension Master change is required.
        '   If the Transfer To Division (use the Division Master to determine this) is not within the same Country, the date of termination equals the Pension Exit Date
        '   and the DB Status Date. The DB Status would equal "X". Create an Alert with the Type of Event equal to Salaried Transfer Out with an Old Value and New Value.
            
            xUpdFlag = False
            If xSalHly = "Hourly" Then
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                
                'No Pension Master record found, add it first
                If rsPen.EOF Then
                    'Pen Master Audit -  Ticket #19954 - begin
                    UpdPenAudit = True
                    Call PenMasterAuditOldValSetup("Blank")
                    toTYPE = "Add"
                    'Pen Master Audit -  Ticket #19954 - end
                    rsPen.Close
                    Call WFCPensionMaster(xEmpNo, "Y", "X", , xYear)
                    'reopen this recordset again
                    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                    SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                    SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                    If rsPen.State <> 0 Then rsPen.Close
                    rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Else
                    'Pen Master Audit -  Ticket #19954 - begin
                    Call PenMasterAuditOldValSetup("CurValues", rsPen)
                    toTYPE = "Change"
                    'Pen Master Audit -  Ticket #19954 - end
                End If
                
                If Not rsPen.EOF Then
                    If IsDate(xUptVal) Then
                        rsPen("PE_EXIT_DATE") = xUptVal
                        rsPen("PE_DB_STATUS_DATE") = xUptVal 'keep as same as salaried emp
                        rsPen("PE_DB_STATUS") = "X"          'keep as same as salaried emp
                        If Not IsMissing(xUserID) Then
                            rsPen("PE_LUSER") = xUserID
                        Else
                            rsPen("PE_LUSER") = glbUserID
                        End If
                        'Ticket #21822 Franks 04/10/2012
                        '"   When transferring from a hourly to salary. The Freeze Rate should be checked
                        rsPen("PE_FREEZE_RATE") = 1
                        rsPen.Update
                        xUpdFlag = True
                    End If
                End If
                'rsPen.Close
            End If
            If xSalHly = "Salaried" Then
                'Get Div Country
                xDivNew = xUptOldVal
                xTmp1 = Get_Division_Name(xDIV, "DV_COUNTRY")       'Transfer Out Div
                xTmp2 = Get_Division_Name(xDivNew, "DV_COUNTRY") 'Transfer In Div
                If Len(xTmp1) > 0 And Len(xTmp2) > 0 Then
                    If xTmp1 = xTmp2 Then 'Same Country
                        'do nothing
                    Else 'Diff Country
                        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                        SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                        SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                        If rsPen.State <> 0 Then rsPen.Close
                        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        
                        'No Pension Master record found, add it first
                        If rsPen.EOF Then
                            'Pen Master Audit -  Ticket #19954 - begin
                            UpdPenAudit = True
                            Call PenMasterAuditOldValSetup("Blank")
                            toTYPE = "Add"
                            'Pen Master Audit -  Ticket #19954 - end
                            rsPen.Close
                            Call WFCPensionMaster(xEmpNo, "Y", "X", , xYear)
                            'reopen this recordset again
                            SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                            SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                            SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                            SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                            If rsPen.State <> 0 Then rsPen.Close
                            rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        Else
                            'Pen Master Audit -  Ticket #19954 - begin
                            Call PenMasterAuditOldValSetup("CurValues", rsPen)
                            toTYPE = "Change"
                            'Pen Master Audit -  Ticket #19954 - end
                        End If
                        
                        If Not rsPen.EOF Then
                            If IsDate(xUptVal) Then
                                rsPen("PE_EXIT_DATE") = xUptVal
                                rsPen("PE_DB_STATUS_DATE") = xUptVal
                                rsPen("PE_DB_STATUS") = "X"
                                rsPen.Update
                                xUpdFlag = True
                                'Pension Alert
                                Call WFCPensionAlerts(xEmpNo, Date, "Salaried Transfer Out", Get_Division_Name(xDivNew), Get_Division_Name(xDIV))
                            End If
                        End If
                        'rsPen.Close
                    End If
                End If
            End If
            
            'Transfer Out Credit Service, Cont. Service - begin
            If xUpdFlag Then
                'Max is 12 months, if new hire in this year, convert the days to months
                xlTmpDateFr = CVDate("Jan 1, " & xYear)
                xTermDate = xUptVal 'Transfer Out date
                xlTmpDateTo = xTermDate
        
                If IsDate(rsEmp("ED_DOH")) Then
                    If CVDate(rsEmp("ED_DOH")) > CVDate(xlTmpDateFr) Then
                        xlTmpDateFr = rsEmp("ED_DOH")
                    End If
                End If
                xContServ = DateDiff("D", xlTmpDateFr, xlTmpDateTo) + 1 'xLastDay)
                If xContServ > 365 Then
                    xContServ = 365
                End If
                If xContServ < 0 Then xContServ = 0
                xContServ = Round((xContServ / (365 / 12)), 2)
                If xContServ >= 11.96 And xContServ <= 12 Then
                    xContServ = 12
                End If
                If xContServ <= 0.07 Then xContServ = 0
                xContServ = Round(xContServ, 1) 'Ticket #23704 Franks 05/06/2013
                
                xCreditService = xContServ
                
                If xSalHly = "Salaried" Then
                    xCreditService = xContServ
                Else
                    'Hourly employees
                    'Check Status change history
                    xLastDay = xTermDate
                    xFirstPenDay = xlTmpDateFr
                    If IsDate(rsEmp("ED_ELIGIBLE")) Then
                        'xPenStartDate = rsEmp("ED_ELIGIBLE")
                        If Not (CVDate(rsEmp("ED_ELIGIBLE")) <= CVDate(xFirstPenDay)) Then 'xFirstDay
                            xFirstPenDay = rsEmp("ED_ELIGIBLE")
                        End If
                    End If
                    xCurStatus = rsEmp("ED_EMP")
                'If xEmpNo = 12352191 Then ' 12109104 Then
                'Debug.Print ""
                'End If
                    Call AddEmpStatus(xEmpNo, xYear, xLastDay, xCurStatus, xFirstPenDay)
                    'Calculate the Credited Service - Begin
                    xlCreditServ = CalcCreditService(xEmpNo, xYear, xUnion, xSection, xDOH, xLastDay)
                    'If xlCreditServ < xCreditService Then
                    If xlCreditServ <= 12 Then
                        xCreditService = xlCreditServ
                    End If
                    'Calculate the Credited Service - End
    
                End If
                
                If xSalHly = "Hourly" Then
                    If Not IsNull(rsPen("PE_NOGC")) Then
                        xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate, rsPen("PE_NOGC"))
                    Else
                        xBenRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate)
                    End If
    
                    If xBenRate > 0 Then
                        If IsNull(rsPen("PE_BENEFIT_RATE")) Then
                            rsPen("PE_BENEFIT_RATE") = xBenRate
                        Else
                            If rsPen("PE_BENEFIT_RATE") = 0 Then
                                rsPen("PE_BENEFIT_RATE") = xBenRate
                            End If
                        End If
                        rsPen("PE_YEAR_AMOUNT") = xCreditService * xBenRate
                    End If
    
                    '''Ticket #20717
                    If xSERYN = "Y" Then
                        If Not IsNull(rsPen("PE_NOGC")) Then
                            xSupRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate, rsPen("PE_NOGC"), "Y")
                        Else
                            xSupRate = GetWFCBenRate(xUnion, rsPen("PE_PENSIONTYPE"), xTermDate, , "Y")
                        End If
        
                        If xSupRate > 0 Then
                            If IsNull(rsPen("PE_SUPPLE_RATE")) Then
                                rsPen("PE_SUPPLE_RATE") = xSupRate
                            Else
                                If rsPen("PE_SUPPLE_RATE") = 0 Then
                                    rsPen("PE_SUPPLE_RATE") = xSupRate
                                End If
                            End If
                        End If
                    End If
    
                End If
                rsPen("PE_CONT_SERV") = xContServ
                rsPen("PE_CREDITED_SERV") = xCreditService
                rsPen.Update
                'Pen Master Audit Ticket #19954
                Call AUDIT_PENSION_MASTER(rsPen)
            End If
            'Transfer Out Credit Service, Cont. Service - end
            
        End If
        
        If xUptType = "Transfer In" Then
            'For hourly employees, automatically create a Pension Master file for the DB Pension.  The pension type will be determined once the Plant and Union code have been entered.  Under the EI Module.
            'For salaried employees, if the Transfer IN Division (use the Division Master to determine this) is still within the same Country as the transfer out, no Pension Master change is required. Otherwise, automatically create a Pension Master file for the DB Pension.  The pension type will be determined once the Plant and Union code have been entered.  Under the EI Module.
            'Pension Master status should be A for active
            xUpdFlag = False
            UpdPenAudit = True 'Ticket #19954
            xTermDate = xUptVal
            If xSalHly = "Hourly" Then
                Call WFCPensionMaster(rsEmp("ED_EMPNBR"), "Y", "A", , xYear, , , , , xTermDate)
                xUpdFlag = True
            End If
            If xSalHly = "Salaried" Then
                xDivNew = xUptOldVal
                If xDIV = xDivNew Then 'Ticket #21677 Franks 03/14/2012
                    'same Div then check if it is Union Transfer
                     If Len(xOldUnion) > 0 Then
                        If Not xUnion = xOldUnion Then
                            If getDBType(xSection, xOldUnion, "HlySal") = "Hourly" Then
                                'union transfer from Hourly to Salaried
                                Call WFCPensionMaster(rsEmp("ED_EMPNBR"), "Y", "A", , xYear, , , , , xTermDate)
                                xUpdFlag = True
                            End If
                        End If
                    End If
                Else
                    'Get Div Country
                    'xDivNew = xUptOldVal
                    xTmp1 = Get_Division_Name(xDIV, "DV_COUNTRY")       'Transfer In Div
                    xTmp2 = Get_Division_Name(xDivNew, "DV_COUNTRY") 'Transfer Out Div
                    If Len(xTmp1) > 0 And Len(xTmp2) > 0 Then
                        If xTmp1 = xTmp2 Then 'Same Country
                            'do nothing
                        Else 'Diff Country
                            Call WFCPensionMaster(rsEmp("ED_EMPNBR"), "Y", "A", , xYear)
                            xUpdFlag = True
                        End If
                    End If
                End If
            End If
            If xUpdFlag Then
                If IsDate(rsEmp("ED_DOB")) Then
                    'Early Retirement Date
                    If Not IsDate(rsEmp("ED_EARLYR")) Then
                        rsEmp("ED_EARLYR") = WFCPenEmpRetireDate(55, rsEmp("ED_DOB"))
                    End If
                    'Normal  Retirement Date
                    If Not IsDate(rsEmp("ED_NORMALR")) Then
                        rsEmp("ED_NORMALR") = WFCPenEmpRetireDate(65, rsEmp("ED_DOB"))
                    End If
                    'Early Retirement Date
                    If Not IsDate(rsEmp("ED_LATESTR")) Then
                        rsEmp("ED_LATESTR") = WFCPenEmpRetireDate(71, rsEmp("ED_DOB"))
                    End If
                    rsEmp.Update
                    'xUptVal Pension transfer in date
                    If IsDate(xUptVal) Then
                        SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo
                        If rsEmOther2.State <> 0 Then rsEmOther2.Close
                        rsEmOther2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                        If rsEmOther2.EOF Then
                            rsEmOther2.AddNew
                            rsEmOther2("ER_EMPNBR") = xEmpNo
                        End If
                        rsEmOther2("ER_PENSIONDATE3") = xUptVal
                        rsEmOther2.Update
                    End If
                End If
            End If
            
        End If 'end of Transfer In
        
        If xUptType = "Position_NOGC" Then
        'Entered a new position for a hourly employee. The employee has a Pension record for the current year.
        '1. if Membership Exit Date was entered, it should do nothing.
        '2. If the Status is one of Termination codes, such as "C", "R", "N", "T",it should do nothing as well.
            If xSalHly = "Hourly" Then
                xCurNOCC = GetNOCC_FromJob(xUptOldVal)
                SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
                SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
                If rsPen.State <> 0 Then rsPen.Close
                rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsPen.EOF Then
                    If IsNull(rsPen("PE_EXIT_DATE")) Then
                        If Not (rsPen("PE_DB_STATUS") = "C" Or rsPen("PE_DB_STATUS") = "R" Or rsPen("PE_DB_STATUS") = "N" Or rsPen("PE_DB_STATUS") = "T") Then
                            rsPen("PE_NOGC") = xCurNOCC
                            rsPen.Update
                        End If
                    End If
                End If
                rsPen.Close
            End If
        End If
        'end of Position_NOGC
    End If
    rsEmp.Close
End Sub

Public Sub WFCPAMaster(xEmpNo, xSIN, xYear, xSalHly, xNewRec, Optional xPENEARN)
'Public Sub WFCPAMaster(xEmpNo, xSIN, xYear, xAnnSal, xPENEARN, xDN02_YTD, xDN43_YTD, XDBPA, xTOTAL_PA, xTERMDATE, xDeemPenFlag As Boolean, xADPDept, Optional xTERM_SEQ)
Dim rsEmp As New ADODB.Recordset
Dim rsPA As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim rsTTemp As New ADODB.Recordset
Dim xTermID
Dim SQLQ As String
Dim I As Double, xTot As Double, xAnnSal As Double, xTOTAL_DBPA
Dim xPE_FLAG, xTtmp, xTStr1 As String, xTStr2 As String

    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If rsEmp.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmp("ED_EMPTYPE")) Then Exit Sub
        If Not rsEmp("ED_EMPTYPE") = "Y" Then
                Exit Sub
        End If
    End If

    SQLQ = "SELECT * FROM HRP_PA_MASTER WHERE PE_SIN = '" & xSIN & "' "
    SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
    SQLQ = SQLQ & "AND PE_HRLYSAL = '" & xSalHly & "' "
    rsPA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsPA.EOF Then 'found it and xNewRec = Y then skip it
        If Not xNewRec = "Y" Then
            Exit Sub
        End If
    End If
    If rsPA.EOF Then
        rsPA.AddNew
        rsPA("PE_COUNTRY") = "CAN"
        rsPA("PE_YEAR_DATE") = xYear
        rsPA("PE_SIN") = xSIN
        rsPA("PE_HRLYSAL") = xSalHly
        rsPA("PE_DN02_YTD") = 0 'xDN02_YTD
        rsPA("PE_DN43_YTD") = 0 'xDN43_YTD
        rsPA("PE_DIV") = rsEmp("ED_DIV")
        rsPA("PE_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
        rsPA("PE_EMPNBR") = xEmpNo
        rsPA("PE_SURNAME") = rsEmp("ED_SURNAME")
        rsPA("PE_FNAME") = rsEmp("ED_FNAME")
        rsPA("PE_SECTION") = rsEmp("ED_SECTION")
        'rsPA("PE_DMD_PENEARN") = xPENEARN 'new record only, use can change it later
    End If

    xAnnSal = 0
    xTOTAL_DBPA = 0
    If xSalHly = "Salaried" Then
        'get Ann Salary
        SQLQ = "SELECT SH_EMPNBR, SH_SALARY FROM HR_SALARY_HISTORY WHERE NOT (SH_CURRENT =0) AND SH_EMPNBR = " & xEmpNo & " "
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsSal.EOF Then
            xAnnSal = rsSal("SH_SALARY")
        End If
        rsPA("PE_ANN_SALARY") = xAnnSal
        If Not IsMissing(xPENEARN) Then
            rsPA("PE_ADP_PENEARN") = xPENEARN
            rsPA("PE_DMD_PENEARN") = xPENEARN
            xTOTAL_DBPA = calDBPA(xPENEARN, xSalHly)
            rsPA("PE_TOTAL_DBPA") = xTOTAL_DBPA
            rsPA("PE_TOTAL_PA") = Round(xTOTAL_DBPA, 0)
        End If
    End If
    
    If xSalHly = "Hourly" Then
        If Not IsMissing(xPENEARN) Then
            rsPA("PE_DMD_PENEARN") = xPENEARN
            xTOTAL_DBPA = calDBPA(xPENEARN, xSalHly)
            rsPA("PE_TOTAL_DBPA") = xTOTAL_DBPA
            rsPA("PE_TOTAL_PA") = Round(xTOTAL_DBPA, 0)
        End If
    End If
    rsPA("PE_LDATE") = Date '
    rsPA("PE_LTIME") = Time$
    rsPA("PE_LUSER") = glbUserID
    rsPA.Update
    rsPA.Close
            
    'create PA Details record - begin
    SQLQ = "SELECT * FROM HRP_PA_DETAILS WHERE PE_SIN = '" & xSIN & "' "
    SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
    SQLQ = SQLQ & "AND PE_HRLYSAL = '" & xSalHly & "' "
    rsPA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsPA.EOF Then 'found it and xNewRec = Y then skip it
        If Not xNewRec = "Y" Then
            Exit Sub
        End If
    End If
    If rsPA.EOF Then
        rsPA.AddNew
        rsPA("PE_COUNTRY") = "CAN"
        rsPA("PE_YEAR_DATE") = xYear
        rsPA("PE_SIN") = xSIN
        rsPA("PE_HRLYSAL") = xSalHly
        rsPA("PE_DN02_YTD") = 0 'xDN02_YTD
        rsPA("PE_DN43_YTD") = 0 'xDN43_YTD
        rsPA("PE_DIV") = rsEmp("ED_DIV")
        rsPA("PE_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
        rsPA("PE_EMPNBR") = xEmpNo
        rsPA("PE_SURNAME") = rsEmp("ED_SURNAME")
        rsPA("PE_FNAME") = rsEmp("ED_FNAME")
        rsPA("PE_SECTION") = rsEmp("ED_SECTION")
    End If 'xAnnSal xTOTAL_DBPA
    
    If Not IsMissing(xPENEARN) Then
        rsPA("PE_ANN_SALARY") = xAnnSal
        rsPA("PE_ADP_PENEARN") = xPENEARN
        rsPA("PE_DBPA") = xTOTAL_DBPA
        rsPA("PE_ADP_TOTAL_PA") = Round(xTOTAL_DBPA, 0)
    End If
    rsPA("PE_LDATE") = Date '
    rsPA("PE_LTIME") = Time$
    rsPA("PE_LUSER") = glbUserID
    rsPA.Update
    rsPA.Close
    'create PA Details record - end
End Sub


Public Sub WFCPensionMasOnType(xEmpNo, xType, xEDate, Optional xPenStatus)
Dim rsEmp As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim rsBen As New ADODB.Recordset
Dim rsHRBeFicia As New ADODB.Recordset
Dim SQLQ As String
Dim xSection, xPenType, xDBStatus, xSalHly
Dim xDOB, xEarlyRet, xNorRet, xLateRet, xDATE, xYear, xSIN, xUnion
Dim xPenExitDate, xOldValue
Dim xPayPeriodAmount

    If Not IsDate(xEDate) Then
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        If IsNull(rsEmp("ED_EMPTYPE")) Then Exit Sub
        If Not rsEmp("ED_EMPTYPE") = "Y" Then
                Exit Sub
        End If
        If IsNull(rsEmp("ED_SIN")) Then Exit Sub
        If IsNull(rsEmp("ED_SECTION")) Then Exit Sub
        If IsNull(rsEmp("ED_DOH")) Then Exit Sub
        If IsNull(rsEmp("ED_DOB")) Then Exit Sub
        If IsNull(rsEmp("ED_ORG")) Then Exit Sub
        'If IsNull(rsEMP("ED_ELIGIBLE")) Then Exit Sub
        
        xSIN = rsEmp("ED_SIN")
        xSection = rsEmp("ED_SECTION")
        xUnion = rsEmp("ED_ORG")
        xYear = Year(xEDate) 'Year(rsEMP("ED_ELIGIBLE"))
        If xType = "DCPP" Then 'Ticket #22964 Franks 12/17/2012
            xPenType = "DC"
        Else
            xPenType = xType 'getDBType(xSection, xUnion, "PenType")
        End If
        xSalHly = getDBType(xSection, xUnion, "HlySal")
        If Len(xPenType) = 0 Then
            Exit Sub
        End If
        If Len(xSalHly) = 0 Then
            Exit Sub
        End If
        If IsMissing(xPenStatus) Then
            xDBStatus = "A" '"H"
        Else
            xDBStatus = xPenStatus
        End If

        'find Benefit record
        SQLQ = "SELECT * FROM HRBENFT "
        SQLQ = SQLQ & " WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & " AND BF_BCODE = '" & xType & "' "
        SQLQ = SQLQ & " AND BF_EDATE = " & Date_SQL(xEDate) & " "
        SQLQ = SQLQ & " ORDER BY BF_BCODE, BF_EDATE "
        If rsBen.State <> 0 Then rsBen.Close
        rsBen.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsBen.EOF Then
            Exit Sub
        End If
        'rsBen.Close
        
        'add pension master record - begin
        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
        If rsPen.State <> 0 Then rsPen.Close
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsPen.EOF Then
            rsPen.AddNew
            rsPen("PE_SIN") = xSIN
            rsPen("PE_YEAR_DATE") = xYear
            rsPen("PE_PENSIONTYPE") = xPenType '"DBS"
            rsPen("PE_HRLYSAL") = xSalHly '"Salaried"
            rsPen("PE_COUNTRY") = "CAN"
            rsPen("PE_DB_STATUS") = xDBStatus
        End If
        rsPen("PE_EMPNBR") = rsEmp("ED_EMPNBR")
        rsPen("PE_SURNAME") = rsEmp("ED_SURNAME")
        rsPen("PE_FNAME") = rsEmp("ED_FNAME")
        rsPen("PE_SECTION") = xSection
        'rsPen("PE_DB_STATUS") = xDBStatus
        'If IsDate(rsEMP("ED_ELIGIBLE")) Then
            rsPen("PE_DB_STATUS_DATE") = xEDate
            rsPen("PE_PEN_ENTRY_DATE") = xEDate
        'End If
        If IsDate(rsBen("BF_CEASEDATE")) Then
            rsPen("PE_EXIT_DATE") = rsBen("BF_CEASEDATE")
        Else
            rsPen("PE_EXIT_DATE") = Null
        End If
        If Not IsNull(rsBen("BF_PPAMT")) Then
            rsPen("PE_MEM_PERCENTAGE") = rsBen("BF_PPAMT")
        End If
        rsPen.Update
        rsPen.Close
        'add pension master record - end

    End If
    rsEmp.Close

End Sub

Public Sub WFCPensionMaster(xEmpNo, Optional xAddBeneficiary, Optional xPenStatus, Optional xUserID, Optional xActionYear, Optional xUptType, Optional xCopyPreRecord, Optional xPType, Optional mSalHly, Optional xPenEDate, Optional xTERM_Seq = 0)
'see Pension Phase II.doc in W:\2008 Projects\Pension\Pension Phase II folder
Dim rsEmp As New ADODB.Recordset
Dim rsPen As New ADODB.Recordset
Dim rsPenPre As New ADODB.Recordset
Dim rsHRBeFicia As New ADODB.Recordset
Dim SQLQ As String
Dim xSection, xPenType, xDBStatus, xSalHly
Dim xDOB, xEarlyRet, xNorRet, xLateRet, xDATE, xYear, xSIN, xUnion
Dim xPenExitDate, xOldValue
Dim xCurNOCC
Dim xPreFlag As Boolean
Dim xPre2Flag As Boolean 'Ticket #23459 Franks 03/26/2013

    If xTERM_Seq > 0 Then ''Ticket #23565 Franks 04/10/2013
        SQLQ = "SELECT * FROM Term_HREMP WHERE TERM_SEQ = " & xTERM_Seq
    Else
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsEmp.EOF Then
        If IsNull(rsEmp("ED_EMPTYPE")) Then Exit Sub
        If Not rsEmp("ED_EMPTYPE") = "Y" Then
            If IsMissing(xUptType) Then
                Exit Sub
            Else
                If Not xUptType = "PenExitDate" Then
                    Exit Sub
                End If
            End If
        End If
        If IsNull(rsEmp("ED_SIN")) Then Exit Sub
        If IsNull(rsEmp("ED_SECTION")) Then Exit Sub
        If IsNull(rsEmp("ED_DOH")) Then Exit Sub
        If IsNull(rsEmp("ED_DOB")) Then Exit Sub
        If IsNull(rsEmp("ED_ORG")) Then Exit Sub
        If IsNull(rsEmp("ED_ELIGIBLE")) Then Exit Sub
        
        xSIN = rsEmp("ED_SIN")
        xSection = rsEmp("ED_SECTION")
        xUnion = rsEmp("ED_ORG")
        If IsMissing(xActionYear) Then
            xYear = Year(rsEmp("ED_ELIGIBLE"))
        Else
            xYear = xActionYear
        End If
            
        If IsMissing(xPType) Then
            xPenType = getDBType(xSection, xUnion, "PenType", rsEmp("ED_DOH")) 'Ticket #26707 Franks 02/25/2015
        Else
            xPenType = xPType
        End If
        If IsMissing(mSalHly) Then
            xSalHly = getDBType(xSection, xUnion, "HlySal")
        Else
            xSalHly = mSalHly
        End If
        If Len(xPenType) = 0 Then
            Exit Sub
        End If
        If Len(xSalHly) = 0 Then
            Exit Sub
        End If
        If IsMissing(xPenStatus) Then
            xDBStatus = "H"
        Else
            xDBStatus = xPenStatus
        End If
        
        If Not IsNull(rsEmp("ED_DOB")) Then
            xDOB = rsEmp("ED_DOB")
        Else
            xDOB = Date
        End If

        'add pension master record - begin
        xCurNOCC = ""
        If xSalHly = "Hourly" Then
            xCurNOCC = GetCurNOCC(xEmpNo)
        End If
        SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
        If Not IsMissing(xUptType) Then
            If xUptType = "Rehire" Then
            'Meeting with Margaret, some employees were Termination - Cash out, then WFC Rehire them back.
            'So info:HR Rehire  should create a new Pension Master record with status "H -New Hire"
            'for this employees even he/she has a Pension Master record in the same year but status is different. - Frank (Jun 9th,, 2010)
                SQLQ = SQLQ & "AND PE_DB_STATUS = '" & xDBStatus & "' "
            End If
            If xUptType = "UsePenStatus" Then
                SQLQ = SQLQ & "AND PE_DB_STATUS = '" & xDBStatus & "'  "
            End If
        End If
        If rsPen.State <> 0 Then rsPen.Close
        rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'Pen Master Audit -  Ticket #19954 - begin
        If rsPen.EOF Then
            UpdPenAudit = True
            Call PenMasterAuditOldValSetup("Blank")
            toTYPE = "Add"
        Else
            Call PenMasterAuditOldValSetup("CurValues", rsPen)
            toTYPE = "Change"
        End If
        'Pen Master Audit -  Ticket #19954 - end
        If rsPen.EOF Then
            'new record - Begin
            rsPen.AddNew
            rsPen("PE_SIN") = xSIN
            rsPen("PE_YEAR_DATE") = xYear
            rsPen("PE_PENSIONTYPE") = xPenType '"DBS"
            rsPen("PE_HRLYSAL") = xSalHly '"Salaried"
            rsPen("PE_COUNTRY") = "CAN"
            rsPen("PE_DB_STATUS") = xDBStatus
            xPreFlag = False
            xPre2Flag = False 'Ticket #23459 Franks 03/26/2013
            If Not IsMissing(xCopyPreRecord) Then
                'If xCopyPreRecord = "Y" Then
                If Left(xCopyPreRecord, 1) = "Y" Then 'Ticket #23459 Franks 03/26/2013 - it can be "Y" or "YY", the 2nd Y will turn on xPre2Flag
                    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
                    SQLQ = SQLQ & "AND PE_YEAR_DATE < " & xYear & " "
                    SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' AND PE_HRLYSAL = '" & xSalHly & "' "
                    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC"
                    If rsPenPre.State <> 0 Then rsPenPre.Close
                    rsPenPre.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If Not rsPenPre.EOF Then
                        xPreFlag = True
                        If Mid(xCopyPreRecord, 2, 1) = "Y" Then 'Ticket #23459 Franks 03/26/2013
                            xPre2Flag = True
                        End If
                    End If
                End If
            End If
            
            If xPreFlag Then
                'Copy previous fields
                rsPen("PE_NOGC") = rsPenPre("PE_NOGC")
                rsPen("PE_EMPNBR") = rsPenPre("PE_EMPNBR")
                rsPen("PE_SURNAME") = rsPenPre("PE_SURNAME")
                rsPen("PE_FNAME") = rsPenPre("PE_FNAME")
                rsPen("PE_SECTION") = rsPenPre("PE_SECTION")
                rsPen("PE_DB_STATUS_DATE") = rsPenPre("PE_DB_STATUS_DATE")
                rsPen("PE_PEN_ENTRY_DATE") = rsPenPre("PE_PEN_ENTRY_DATE")
                rsPen("PE_EXIT_DATE") = rsPenPre("PE_EXIT_DATE")
                If xPre2Flag Then 'Ticket #23459 Franks 03/26/2013
                    rsPen("PE_ANNDEFERRED") = rsPenPre("PE_ANNDEFERRED")
                    rsPen("PE_FORMPENSION") = rsPenPre("PE_FORMPENSION")
                    rsPen("PE_GUARANTEEPERIOD") = rsPenPre("PE_GUARANTEEPERIOD")
                    rsPen("PE_JOINTPERCENTAGE") = rsPenPre("PE_JOINTPERCENTAGE")
                End If
                rsPen("PE_MEM_ACCOUNT") = rsPenPre("PE_MEM_ACCOUNT")
            Else
                'Not Copy previous fields
                If Len(xCurNOCC) > 0 Then 'new record only
                    rsPen("PE_NOGC") = xCurNOCC
                End If
                '-------
                rsPen("PE_EMPNBR") = rsEmp("ED_EMPNBR")
                rsPen("PE_SURNAME") = rsEmp("ED_SURNAME")
                rsPen("PE_FNAME") = rsEmp("ED_FNAME")
                rsPen("PE_SECTION") = xSection
                If Not IsMissing(xPenEDate) Then 'Ticket #21786 Franks 03/26/2012
                    If IsDate(xPenEDate) Then
                        rsPen("PE_DB_STATUS_DATE") = xPenEDate
                        rsPen("PE_PEN_ENTRY_DATE") = xPenEDate
                    End If
                Else
                    If IsDate(rsEmp("ED_ELIGIBLE")) Then
                        rsPen("PE_DB_STATUS_DATE") = rsEmp("ED_ELIGIBLE")
                        rsPen("PE_PEN_ENTRY_DATE") = rsEmp("ED_ELIGIBLE")
                    End If
                End If
                rsPen("PE_EXIT_DATE") = Null
            End If
            'new record - end
        Else
            'existing record
            If Not IsNull(rsPen("PE_DB_STATUS")) Then
                If Not (rsPen("PE_DB_STATUS") = "C" Or rsPen("PE_DB_STATUS") = "R") Then
                    rsPen("PE_DB_STATUS") = xDBStatus
                End If
            End If
        End If

        If IsMissing(xUserID) Then
            rsPen("PE_LUSER") = glbUserID
        Else
            rsPen("PE_LUSER") = xUserID
        End If
        rsPen("PE_LDATE") = Date 'Ticket #21963 Franks 04/30/2012
        rsPen("PE_LTIME") = Time$
        rsPen.Update
        'Pen Master Audit Ticket #19954
        If UpdPenAudDirect Then
            Call AUDIT_PENSION_MASTER(rsPen)
        End If
        rsPen.Close
        'add pension master record - end
        
        If Not IsMissing(xAddBeneficiary) Then
            If xAddBeneficiary = "Y" Then
                ''rsHRBeFicia
                'SQLQ = "SELECT * FROM HRBENS "
                'SQLQ = SQLQ & " WHERE BD_EMPNBR = " & xEmpNo & " "
                ''SQLQ = SQLQ & " AND LEFT(BD_BCODE,3) = 'LIF' "
                'SQLQ = SQLQ & " AND BD_BCODE = 'DB' "
                'SQLQ = SQLQ & " ORDER BY BD_BCODE, BD_LDATE "
                'rsHRBeFicia.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                'If Not rsHRBeFicia.EOF Then
                '    Call WFCPensionBeneficiary(xEmpNo, rsHRBeFicia("BD_BCODE"))
                'End If
                'rsHRBeFicia.Close
                Call WFCPensionBeneficiary(xEmpNo, "DB")
            End If
        End If


    End If
    rsEmp.Close
End Sub

Public Function GetNOCC_FromJob(xJobCode)
Dim rsETemp As New ADODB.Recordset
Dim SQLQ As String
Dim xRetVal, xTJOB
    xRetVal = ""
    SQLQ = "SELECT JB_CODE,JB_FEDGRP FROM HRJOB WHERE JB_CODE = '" & xJobCode & "' "
    rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsETemp.EOF Then
        If Not IsNull(rsETemp("JB_FEDGRP")) Then
            xRetVal = rsETemp("JB_FEDGRP")
        End If
    End If
    rsETemp.Close
    GetNOCC_FromJob = xRetVal
End Function
Public Function GetCurNOCC(xEmpnbr) ', xLastDay, Optional xTermID)
Dim rsETemp As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xTermID
Dim xRetVal, xTJOB
    xRetVal = ""
    xTJOB = ""
    xTermID = 0
    SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & xEmpnbr & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        SQLQ = "SELECT ED_EMPNBR,TERM_SEQ FROM TERM_HREMP WHERE ED_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "ORDER BY TERM_SEQ DESC "
        If rsEmp.State <> 0 Then rsEmp.Close
        rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmp.EOF Then
            xTermID = rsEmp("TERM_SEQ")
        End If
    End If
    
    If xTermID = 0 Then 'active emp
        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND NOT (JH_CURRENT = 0) "
        'SQLQ = SQLQ & "AND JH_SDATE <= " & Date_SQL(xLastDay) & " "
        SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
        rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsETemp.EOF Then
            xTJOB = rsETemp("JH_JOB")
        End If
        rsETemp.Close
        If Len(xTJOB) > 0 Then
            SQLQ = "SELECT JB_CODE,JB_FEDGRP FROM HRJOB WHERE JB_CODE = '" & xTJOB & "' "
            rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsETemp.EOF Then
                If Not IsNull(rsETemp("JB_FEDGRP")) Then
                    xRetVal = rsETemp("JB_FEDGRP")
                End If
            End If
            rsETemp.Close
        End If
    Else 'Term
        SQLQ = "SELECT * FROM TERM_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTermID & " "
        SQLQ = SQLQ & "AND NOT (JH_CURRENT = 0) "
        'SQLQ = SQLQ & "AND JH_SDATE <= " & Date_SQL(xLastDay) & " "
        SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
        rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsETemp.EOF Then
            xTJOB = rsETemp("JH_JOB")
        End If
        rsETemp.Close
        If Len(xTJOB) > 0 Then
            SQLQ = "SELECT JB_CODE,JB_FEDGRP FROM HRJOB WHERE JB_CODE = '" & xTJOB & "' "
            rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsETemp.EOF Then
                If Not IsNull(rsETemp("JB_FEDGRP")) Then
                    xRetVal = rsETemp("JB_FEDGRP")
                End If
            End If
            rsETemp.Close
        End If
    End If
    GetCurNOCC = xRetVal
End Function

Public Function getDBType(xPlantCode, xUnion, xImp, Optional xDOH) 'Ticket #26707 Franks 02/25/2015
Dim rsPenTypeMatrix As New ADODB.Recordset
Dim xRetVal As String
Dim SQLQ As String
    xRetVal = ""
    If IsNull(xUnion) Then xUnion = ""
    If IsNull(xPlantCode) Then xPlantCode = ""
    SQLQ = "SELECT * FROM HRP_PENTYPE_MATRIX WHERE PE_SECTION = '" & xPlantCode & "' "
    SQLQ = SQLQ & "AND PE_UNION = '" & xUnion & "' "
    rsPenTypeMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsPenTypeMatrix.EOF Then
        If xImp = "PenType" Then
            If Not IsNull(rsPenTypeMatrix("PE_PENSIONTYPE")) Then
                xRetVal = rsPenTypeMatrix("PE_PENSIONTYPE")
                'Ticket #26707 Franks 02/25/2015 - begin
                If xPlantCode = "BLEN" Then
                    If xRetVal = "DBBLEN" Then
                        If Not IsMissing(xDOH) Then
                            If IsDate(xDOH) Then
                                If CVDate(xDOH) >= CVDate("Jan 1, 2013") Then
                                    xRetVal = "DBDCBLEN"
                                End If
                            End If
                        End If
                    End If
                End If
                'Ticket #26707 Franks 02/25/2015 - end
            End If
        End If
        If xImp = "HlySal" Then
            If Not IsNull(rsPenTypeMatrix("PE_SALARYHOURLY")) Then
                xRetVal = rsPenTypeMatrix("PE_SALARYHOURLY")
            End If
            If xRetVal = "H" Then
                xRetVal = "Hourly"
            End If
            If xRetVal = "S" Then
                xRetVal = "Salaried"
            End If
        End If
    End If
    rsPenTypeMatrix.Close
    getDBType = xRetVal

'    If xPlantCode = "MORV" Then
'        xretVal = "DBKITCH"
'    ElseIf xPlantCode = "WHBY" Then
'        xretVal = "DBWHIT"
'    Else
'        xretVal = "DB" & xPlantCode
'    End If
'    getDBType = xretVal
End Function


Public Function WFCPenEmpRetireDate(xYear, xDOB)
Dim xDATE As Date
    'xDATE = DateAdd("YYYY", 55, rsEmp("ED_DOB"))
    xDATE = DateAdd("YYYY", xYear, xDOB)
    xDATE = CVDate(MonthName(month(xDATE)) & " 1," & Year(xDATE))
    xDATE = DateAdd("M", 1, xDATE)
    WFCPenEmpRetireDate = xDATE
End Function

Public Function GetPensionType(xSecCode, xUnionCode)
Dim rsSH As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    If Len(xSecCode) > 0 And Len(xUnionCode) > 0 Then
        SQLQ = "SELECT * FROM HRP_PENTYPE_MATRIX WHERE PE_SECTION = '" & xSecCode & "' "
        SQLQ = SQLQ & "AND PE_UNION = '" & xUnionCode & "' "
        rsSH.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsSH.EOF Then
            If Not IsNull(rsSH("PE_PENSIONTYPE")) Then
                    retVal = rsSH("PE_PENSIONTYPE")
            End If
        End If
        rsSH.Close
    End If
    GetPensionType = retVal
End Function
Public Function GetSalHourly(xSecCode, xUnionCode)
Dim rsSH As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    If Len(xSecCode) > 0 And Len(xUnionCode) > 0 Then
        SQLQ = "SELECT * FROM HRP_PENTYPE_MATRIX WHERE PE_SECTION = '" & xSecCode & "' "
        SQLQ = SQLQ & "AND PE_UNION = '" & xUnionCode & "' "
        rsSH.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsSH.EOF Then
            If Not IsNull(rsSH("PE_SALARYHOURLY")) Then
                If rsSH("PE_SALARYHOURLY") = "H" Then
                    retVal = "Hourly"
                End If
                If rsSH("PE_SALARYHOURLY") = "S" Then
                    retVal = "Salaried"
                End If
            End If
        End If
        rsSH.Close
    End If
    GetSalHourly = retVal
End Function

Public Function GetWFCBenRate(xUnion, xPenType, xLastDay, Optional xNOGC, Optional xSERYN)
Dim rsBenRate As New ADODB.Recordset
Dim SQLQ As String
Dim SERYN As String
Dim xRetVal
        SERYN = "N"
        If Not IsMissing(xSERYN) Then
            SERYN = xSERYN
        End If
        SQLQ = "SELECT * FROM HRP_HOURLY_PEN_RATES WHERE (1=1) "
        SQLQ = SQLQ & "AND PE_UNION = '" & xUnion & "' "
        SQLQ = SQLQ & "AND PE_PENSIONTYPE = '" & xPenType & "' "
        SQLQ = SQLQ & "AND PE_EFFEC_DATE <= " & Date_SQL(xLastDay) & " "
        If Not IsMissing(xNOGC) Then
            SQLQ = SQLQ & "AND PE_NOGC = '" & xNOGC & "' "
        End If
        SQLQ = SQLQ & "ORDER BY PE_EFFEC_DATE DESC "
        rsBenRate.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xRetVal = 0
        If Not rsBenRate.EOF Then
            If SERYN = "Y" Then
                If Not IsNull(rsBenRate("PE_SUPPLE_RATE")) Then
                    xRetVal = rsBenRate("PE_SUPPLE_RATE")
                End If
            Else
                If Not IsNull(rsBenRate("PE_BENEFIT_RATE")) Then
                    xRetVal = rsBenRate("PE_BENEFIT_RATE")
                End If
            End If
        End If
        rsBenRate.Close
        GetWFCBenRate = xRetVal
End Function

Public Function WFCPensionEligible(EmpNbr, Optional xTermSEQ) As Boolean
    Dim rsEmp As New ADODB.Recordset, rsTABL As New ADODB.Recordset
    
    rsEmp.Open "SELECT ED_EMPTYPE FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001
    If rsEmp.EOF Then
        WFCPensionEligible = False
        rsEmp.Close
        Exit Function
    End If
    If UCase(rsEmp("ED_EMPTYPE")) = "Y" Then
        WFCPensionEligible = True
    End If
    rsEmp.Close
End Function

Public Sub AddEmpStatus(xEmpnbr, xYear, xLastDay, xCurStatus, xFirstPenDay, Optional xTermID)
Dim rsEStatus As New ADODB.Recordset
Dim rsEWrk As New ADODB.Recordset
Dim rsETerm As New ADODB.Recordset
Dim rsETrmEmp As New ADODB.Recordset
Dim SQLQ As String
Dim xTmpStatus As String
Dim xTpDate
Dim xTmEndDate
Dim xTmESLastCode
    
'If xEmpnbr = 12352484 Then ' 12030853 Then'12208478 '12001443
'Debug.Print ""
'End If
    
SQLQ = "DELETE FROM HREMPHIS_WRK WHERE EE_WRKEMP = '" & glbUserID & "' "
gdbAdoIhr001.Execute SQLQ

    'Check Status History
If IsMissing(xTermID) Then 'active emp
    
    SQLQ = "SELECT * FROM HREMPHIS WHERE EE_EMPNBR = " & xEmpnbr & " "
    SQLQ = SQLQ & "AND NOT (EE_OLDSTAT IS NULL) "
    SQLQ = SQLQ & "AND NOT (EE_NEWSTAT IS NULL) " 'xFirstDay ,xLastDay
    SQLQ = SQLQ & "AND NOT (EE_CHGDATE IS NULL) "
    SQLQ = SQLQ & "AND EE_CHGDATE >= " & Date_SQL(xFirstPenDay) & " "
    SQLQ = SQLQ & "AND EE_CHGDATE <= " & Date_SQL(xLastDay) & " "
    SQLQ = SQLQ & "ORDER BY EE_CHGDATE,EE_LDATE "
    rsEStatus.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEStatus.EOF Then
        'Find the last statu Code before 12/31/yyyy in Emp Histroy
        xTmESLastCode = GetLastStatusCode(xEmpnbr, xLastDay, "Y")
        If Len(xTmESLastCode) > 0 Then
            xCurStatus = xTmESLastCode
        Else
            'Debug.Print ""
        End If
        
        'No Status change, same emp status for whole year
        SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND EE_OLDVALUE = '" & xCurStatus & "' "
        SQLQ = SQLQ & "AND EE_CHGDATE = " & Date_SQL(xFirstPenDay) & " "
        SQLQ = SQLQ & "AND EE_DOT = " & Date_SQL(xLastDay) & " "
        SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
        If rsEWrk.State <> 0 Then rsEWrk.Close
        rsEWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEWrk.EOF Then
                rsEWrk.AddNew
                rsEWrk("EE_EMPNBR") = xEmpnbr
                rsEWrk("EE_OLDVALUE") = xCurStatus
                rsEWrk("EE_CHGDATE") = xFirstPenDay
                rsEWrk("EE_DOT") = xLastDay                     'End Date
                rsEWrk("EE_WRKEMP") = glbUserID
                rsEWrk.Update
        End If
    Else
        xTpDate = xFirstPenDay
        Do While Not rsEStatus.EOF
            If Not (rsEStatus("EE_OLDSTAT") = rsEStatus("EE_NEWSTAT")) Then
                xTmEndDate = (DateAdd("D", -1, rsEStatus("EE_CHGDATE")))
                If CVDate(xTmEndDate) < CVDate(xTpDate) Then
                    xTmEndDate = xTpDate
                End If
                SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & xEmpnbr & " "
                SQLQ = SQLQ & "AND EE_OLDVALUE = '" & rsEStatus("EE_OLDSTAT") & "' "
                SQLQ = SQLQ & "AND EE_CHGDATE = " & Date_SQL(xTpDate) & " "
                SQLQ = SQLQ & "AND EE_DOT = " & Date_SQL(xTmEndDate) & " "
                SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
                If rsEWrk.State <> 0 Then rsEWrk.Close
                rsEWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If rsEWrk.EOF Then
                    rsEWrk.AddNew
                    rsEWrk("EE_EMPNBR") = xEmpnbr
                    rsEWrk("EE_OLDVALUE") = rsEStatus("EE_OLDSTAT")
                    rsEWrk("EE_CHGDATE") = xTpDate
                    rsEWrk("EE_DOT") = xTmEndDate                 'End Date
                    rsEWrk("EE_WRKEMP") = glbUserID
                    rsEWrk.Update
                End If
            End If
NextRec:
            xTpDate = rsEStatus("EE_CHGDATE")
            xCurStatus = rsEStatus("EE_NEWSTAT")
            rsEStatus.MoveNext
        Loop
        'Add the current status code
        If CVDate(xTpDate) < CVDate(xLastDay) Then '
            SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & xEmpnbr & " "
            SQLQ = SQLQ & "AND EE_OLDVALUE = '" & xCurStatus & "' "
            SQLQ = SQLQ & "AND EE_CHGDATE = " & Date_SQL(xTpDate) & " "
            SQLQ = SQLQ & "AND EE_DOT = " & Date_SQL(xLastDay) & " "
            SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
            If rsEWrk.State <> 0 Then rsEWrk.Close
            rsEWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsEWrk.EOF Then
                rsEWrk.AddNew
                rsEWrk("EE_EMPNBR") = xEmpnbr
                rsEWrk("EE_OLDVALUE") = xCurStatus
                rsEWrk("EE_CHGDATE") = xTpDate
                rsEWrk("EE_DOT") = xLastDay                     'End Date
                rsEWrk("EE_WRKEMP") = glbUserID
                rsEWrk.Update
            End If
        End If
    End If
    rsEStatus.Close
    
Else 'term
    SQLQ = "SELECT * FROM TERM_HREMPHIS WHERE EE_EMPNBR = " & xEmpnbr & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermID & " "
    SQLQ = SQLQ & "AND NOT (EE_OLDSTAT IS NULL) "
    SQLQ = SQLQ & "AND NOT (EE_NEWSTAT IS NULL) " 'xFirstDay ,xLastDay
    SQLQ = SQLQ & "AND NOT (EE_CHGDATE IS NULL) "
    SQLQ = SQLQ & "AND EE_CHGDATE >= " & Date_SQL(xFirstPenDay) & " "
    SQLQ = SQLQ & "AND EE_CHGDATE <= " & Date_SQL(xLastDay) & " "
    SQLQ = SQLQ & "ORDER BY EE_CHGDATE "
    rsEStatus.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEStatus.EOF Then
        
        xTmESLastCode = GetLastStatusCode(xEmpnbr, xLastDay, "Y", xTermID)
        If Len(xTmESLastCode) > 0 Then
            xCurStatus = xTmESLastCode
        Else
            'Debug.Print ""
        End If
        'No Status change, same emp status for whole year
        SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND EE_OLDVALUE = '" & xCurStatus & "' "
        SQLQ = SQLQ & "AND EE_CHGDATE = " & Date_SQL(xFirstPenDay) & " "
        SQLQ = SQLQ & "AND EE_DOT = " & Date_SQL(xLastDay) & " "
        SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
        If rsEWrk.State <> 0 Then rsEWrk.Close
        rsEWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEWrk.EOF Then
                rsEWrk.AddNew
                rsEWrk("EE_EMPNBR") = xEmpnbr
                rsEWrk("EE_OLDVALUE") = xCurStatus
                rsEWrk("EE_CHGDATE") = xFirstPenDay
                rsEWrk("EE_DOT") = xLastDay                     'End Date
                rsEWrk("EE_WRKEMP") = glbUserID
                rsEWrk("TERM_SEQ") = xTermID
                rsEWrk.Update
        End If
    Else
        xTpDate = xFirstPenDay
        Do While Not rsEStatus.EOF
            If Not (rsEStatus("EE_OLDSTAT") = rsEStatus("EE_NEWSTAT")) Then
                xTmEndDate = (DateAdd("D", -1, rsEStatus("EE_CHGDATE")))
                If CVDate(xTmEndDate) < CVDate(xTpDate) Then
                    xTmEndDate = xTpDate
                End If
                SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & xEmpnbr & " "
                SQLQ = SQLQ & "AND EE_OLDVALUE = '" & rsEStatus("EE_OLDSTAT") & "' "
                SQLQ = SQLQ & "AND EE_CHGDATE = " & Date_SQL(xTpDate) & " "
                SQLQ = SQLQ & "AND EE_DOT = " & Date_SQL(xTmEndDate) & " "
                SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
                If rsEWrk.State <> 0 Then rsEWrk.Close
                rsEWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If rsEWrk.EOF Then
                    rsEWrk.AddNew
                    rsEWrk("EE_EMPNBR") = xEmpnbr
                    rsEWrk("EE_OLDVALUE") = rsEStatus("EE_OLDSTAT")
                    rsEWrk("EE_CHGDATE") = xTpDate
                    rsEWrk("EE_DOT") = xTmEndDate                   'End Date
                    rsEWrk("EE_WRKEMP") = glbUserID
                    rsEWrk("TERM_SEQ") = xTermID
                    rsEWrk.Update
                End If
            End If
NextRecT:
            xTpDate = rsEStatus("EE_CHGDATE")
            xCurStatus = rsEStatus("EE_NEWSTAT")
            rsEStatus.MoveNext
        Loop
        'Add the current status code
        If CVDate(xTpDate) < CVDate(xLastDay) Then '
            SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & xEmpnbr & " "
            SQLQ = SQLQ & "AND EE_OLDVALUE = '" & xCurStatus & "' "
            SQLQ = SQLQ & "AND EE_CHGDATE = " & Date_SQL(xTpDate) & " "
            SQLQ = SQLQ & "AND EE_DOT = " & Date_SQL(xLastDay) & " "
            SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
            If rsEWrk.State <> 0 Then rsEWrk.Close
            rsEWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsEWrk.EOF Then
                rsEWrk.AddNew
                rsEWrk("EE_EMPNBR") = xEmpnbr
                rsEWrk("EE_OLDVALUE") = xCurStatus
                rsEWrk("EE_CHGDATE") = xTpDate
                rsEWrk("EE_DOT") = xLastDay                     'End Date
                rsEWrk("EE_WRKEMP") = glbUserID
                rsEWrk("TERM_SEQ") = xTermID
                rsEWrk.Update
            End If
        End If
    End If
    rsEStatus.Close
    
End If

End Sub

Public Function GetLastStatusCode(xEmpnbr, xDATE, xNew, Optional xTermID)
Dim rsETemp As New ADODB.Recordset
Dim SQLQ As String
Dim xRetVal
    xRetVal = ""
    If IsMissing(xTermID) Then 'active
        SQLQ = "SELECT * FROM HREMPHIS WHERE EE_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND NOT (EE_OLDSTAT IS NULL) "
        SQLQ = SQLQ & "AND NOT (EE_NEWSTAT IS NULL) "
        SQLQ = SQLQ & "AND EE_CHGDATE <= " & Date_SQL(xDATE) & " " '<
        SQLQ = SQLQ & "ORDER BY EE_CHGDATE DESC "
        rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsETemp.EOF Then
            If xNew = "Y" Then
                xRetVal = rsETemp("EE_NEWSTAT")
            End If
            If xNew = "N" Then
                xRetVal = rsETemp("EE_OLDSTAT")
            End If
        End If
        rsETemp.Close
    Else 'Term
        SQLQ = "SELECT * FROM Term_HREMPHIS WHERE EE_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTermID & " "
        SQLQ = SQLQ & "AND NOT (EE_OLDSTAT IS NULL) "
        SQLQ = SQLQ & "AND NOT (EE_NEWSTAT IS NULL) "
        SQLQ = SQLQ & "AND EE_CHGDATE <= " & Date_SQL(xDATE) & " " '<
        SQLQ = SQLQ & "ORDER BY EE_CHGDATE DESC"
        rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsETemp.EOF Then
            If xNew = "Y" Then
                xRetVal = rsETemp("EE_NEWSTAT")
            End If
            If xNew = "N" Then
                xRetVal = rsETemp("EE_OLDSTAT")
            End If
        End If
        rsETemp.Close
    
    End If
    GetLastStatusCode = xRetVal
End Function

Public Sub GetCreditServiceRules(xtUnion, xtSection, xtCurStatus)
Dim rsETemp As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HRP_CREDITED_SERVICE_RULES "
    SQLQ = SQLQ & "WHERE PE_UNION = '" & xtUnion & "' "
    SQLQ = SQLQ & "AND PE_SECTION = '" & xtSection & "' "
    SQLQ = SQLQ & "AND PE_EMPSTATUS = '" & xtCurStatus & "' "
    rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xBreakInService = "": xLeaveMonth = 0: xExpYears = 0: x15DayLogic = True
    If Not rsETemp.EOF Then
        xBreakInService = UCase(rsETemp("PE_BREAK_IN_SERVICE"))
        If Not IsNull(rsETemp("PE_LEAVE_MONTH")) Then
            xLeaveMonth = rsETemp("PE_LEAVE_MONTH")
        End If
        If Not IsNull(rsETemp("PE_EXPERIENCE_YEARS")) Then
            xExpYears = rsETemp("PE_EXPERIENCE_YEARS")
        End If
        If Not IsNull(rsETemp("PE_15DAYS_RULE")) Then
            x15DayLogic = rsETemp("PE_15DAYS_RULE")
        End If
    End If
    rsETemp.Close

End Sub

Public Function CalcCreditService(xEmpnbr, xYear, xEmpUnion, xSection, xDOH, xLastDay, Optional xTermID)
Dim rsEWrk As New ADODB.Recordset
Dim SQLQ As String
Dim xTmpStatus As String
Dim xtmpdate, xTmpDat2, xTmpStartDate, xTmpEndDate
Dim xTmpVal, xTmpVa2
Dim xTmpMaxMth, xTmpDays
Dim xTmFirstDay, xTmLastDay
Dim xRetVal
Dim xLastNewCode
Dim locCSFlag As Boolean

    locErrorNo = 0
    xRetVal = 0
    SQLQ = "SELECT * FROM HREMPHIS_WRK WHERE EE_EMPNBR = " & xEmpnbr & " "
    If IsMissing(xTermID) Then
        SQLQ = SQLQ & "AND TERM_SEQ IS NULL "
    Else
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTermID & " "
    End If
    SQLQ = SQLQ & "AND EE_WRKEMP = '" & glbUserID & "' "
    'SQLQ = SQLQ & "AND NOT (EE_OLDVALUE IS NULL) "
    SQLQ = SQLQ & "ORDER BY EE_CHGDATE "
    rsEWrk.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    Do While Not rsEWrk.EOF
    
'If xEmpnbr = 12109098 Then '12208283 '12030853 '12030120 '12031303
' Debug.Print ""
'End If

        'New condition
        If Not Year(rsEWrk("EE_DOT")) = xYear Then
            GoTo Next_Rec
        End If
        
        xTmpStatus = rsEWrk("EE_OLDVALUE")
        ''check the last NewValue, if they are same keep it; otherwise use last NewValue as xTmpStatus
        ''This action can reduce a bad data in emp history, such as
        ''06/23/2008 EE_OLDVALUE is "TLAY" and EE_NEWVALUE is "ACT"
        ''07/21/2008 EE_OLDVALUE is "TLAY" and EE_NEWVALUE is "ACT"
        'If IsMissing(xTermID) Then
        '    xLastNewCode = GetLastStatusCode(xEmpnbr, rsEWrk("EE_CHGDATE"), "Y")
        'Else
        '    xLastNewCode = GetLastStatusCode(xEmpnbr, rsEWrk("EE_CHGDATE"), "Y", xTermID)
        'End If
        'If Len(xLastNewCode) > 0 Then
        '    If Not (xLastNewCode = xTmpStatus) Then
        '        xTmpStatus = xLastNewCode
        '    End If
        'End If
        
        'Get xBreakInService ,xLeaveMonth, xExpYears
        Call GetCreditServiceRules(xEmpUnion, xSection, xTmpStatus)



        'Break In Service Codes:
        'N - Always accrue Credited Service
        'Y - Never accrue Credited Service
        'U - Use the Rules to calculate Credited Service
        'For eligible codes, such as ACT, FMLA, MAT, ..., give all credit
        If xBreakInService = "N" Then
            'xTmpDays = DateDiff("D", rsEWrk("EE_CHGDATE"), rsEWrk("EE_DOT"))
            'xTmpDays = xTmpDays / (365 / 12)
            'xTmpDays = Round(xTmpDays, 2)
            'xRetVal = xRetVal + xTmpDays
            
            'new - begin
            xTmFirstDay = rsEWrk("EE_CHGDATE")
            If CVDate(xTmFirstDay) < CVDate("Jan 1, " & xYear) Then
                xTmFirstDay = CVDate("Jan 1, " & xYear)
            End If
            xTmLastDay = rsEWrk("EE_DOT")
            'new -end
            
            'xTmpDays = DateDiff("M", rsEWrk("EE_CHGDATE"), rsEWrk("EE_DOT")) - 1
            xTmpDays = DateDiff("M", xTmFirstDay, rsEWrk("EE_DOT")) - 1
            If x15DayLogic Then
                'If Day(rsEWrk("EE_CHGDATE")) <= 15 Then 'Begin Month
                If Day(xTmFirstDay) <= 15 Then 'Begin Month
                    xTmpDays = xTmpDays + 1
                End If
                If Day(rsEWrk("EE_DOT")) >= 15 Then 'End Month  > 15
                    xTmpDays = xTmpDays + 1
                End If
            Else
                xTmpDays = xTmpDays + 2
            End If
            If xTmpDays < 0 Then xTmpDays = 0
            xRetVal = xRetVal + xTmpDays
        End If
        'for non eligible codes, no credit
        If xBreakInService = "Y" Then
            'xRetVal = 0
        End If
        
        If xBreakInService = "U" Then ' And xLeaveMonth > 0 Then
            xTmpDays = 0 'default to 0
            'Get last status code (TALY) date from Emp His table
            'xTmpDate = GetStatusDate(xEmpnbr, "TLAY", xFirstDay)
            xTmFirstDay = rsEWrk("EE_CHGDATE")
            xTmLastDay = rsEWrk("EE_DOT")
            'xTmpDate is Status start Date
            If IsMissing(xTermID) Then
                xtmpdate = GetStatusDate(xEmpnbr, xTmpStatus, xTmFirstDay, "Y")
            Else
                xtmpdate = GetStatusDate(xEmpnbr, xTmpStatus, xTmFirstDay, "Y", xTermID)
            End If
            xTmpStartDate = xtmpdate
            'if Status End Date can be found in this period
            If IsMissing(xTermID) Then
                xTmpEndDate = GetStatusDate(xEmpnbr, xTmpStatus, DateAdd("D", 1, xTmLastDay), "N")
            Else
                xTmpEndDate = GetStatusDate(xEmpnbr, xTmpStatus, DateAdd("D", 1, xTmLastDay), "N", xTermID)
            End If
            
            If IsDate(xtmpdate) Then
                'Get the Leave Month
                'xTmpVal = GetLeaveMonth(xEmpUnion, xYear)
                'If IsNumeric(xTmpVal) Then
                xTmpMaxMth = DateDiff("M", xTmFirstDay, xTmLastDay) - 1
                
                If x15DayLogic Then
                    If Day(xTmFirstDay) <= 15 Then 'Begin Month < 15
                        xTmpMaxMth = xTmpMaxMth + 1
                    End If
                    If Day(xTmLastDay) >= 15 Then 'End Month  > 15
                        xTmpMaxMth = xTmpMaxMth + 1
                    End If
                Else
                    xTmpMaxMth = xTmpMaxMth + 2
                End If
            
                If xTmpMaxMth > 12 Then xTmpMaxMth = 12
                If IsNumeric(xLeaveMonth) Then
                    If xLeaveMonth > 0 Then
                        'xTmpDate = DateAdd("M", xLeaveMonth, xTmpDate)
                        'xTmpDays = DateDiff("M", xTmFirstDay, xTmpDate)
                        'If x15DayLogic Then
                        '    If Day(xTmpDate) < 15 Then
                        '        xTmpDays = xTmpDays + 1
                        '    End If
                        'Else
                        '    xTmpDays = xTmpDays + 1
                        'End If
                        xtmpdate = DateAdd("M", xLeaveMonth, xtmpdate)
                        'xTmpDays = DateDiff("M", xTmFirstDay, xTmpDate) - 1
                        If x15DayLogic Then
                            xTmpDays = DateDiff("M", xTmFirstDay, xtmpdate) - 1
                            If Day(xTmFirstDay) <= 15 Then 'Begin Month
                                xTmpDays = xTmpDays + 1
                            End If
                            If Day(xtmpdate) >= 15 Then
                                xTmpDays = xTmpDays + 1
                            End If
                        Else
                            'get leave month
                            xTmpDays = DateDiff("M", xTmFirstDay, xtmpdate)
                            If xTmpDays < 0 Then
                                xTmpDays = 0
                            End If
                            '==============None 15DaysLogic Begin ====
                            'This section will add the missed CS for Break In Service 'N' Code (ACT)
                            'if they switch each oter, such as TLAY to ACT, ACT to TLAY
                            'xTmpDays = xTmpDays + 2
                            'If status start date = Frist date of this period then + 1 month - Non 15Days Rule
                            If CVDate(xTmpStartDate) = CVDate(xTmFirstDay) Then
                                'Old Code should be Break In Service "N"
                                'Date < then 15th , TLAY start Date is ACT end Date
                                'then + 1

                                'If IsMissing(xTermID) Then
                                '    locCSFlag = GetFromOrToStatus(xEmpnbr, xTmpStatus, xTmpStartDate, "Y")
                                'Else
                                '    locCSFlag = GetFromOrToStatus(xEmpnbr, xTmpStatus, xTmpStartDate, "Y", xTermID)
                                'End If
                                'If locCSFlag Then
                                    If Day(xTmpStartDate) < 15 Then
                                        xTmpDays = xTmpDays + 1
                                    End If
                                'End If
                            End If
                            'If status start date = Frist date of this period then + 1 month - Non 15Days Rule
                            If IsDate(xTmpEndDate) Then
                                If CVDate(xTmpEndDate) = CVDate(DateAdd("D", 1, xTmLastDay)) Then
                                'Old Code should be Break In Service "N"
                                'Date >= then 15th , TLAY End Date is ACT Start Date
                                'then + 1
                                
                                    'If IsMissing(xTermID) Then
                                    '    locCSFlag = GetFromOrToStatus(xEmpnbr, xTmpStatus, xTmpEndDate, "N")
                                    'Else
                                    '    locCSFlag = GetFromOrToStatus(xEmpnbr, xTmpStatus, xTmpEndDate, "N", xTermID)
                                    'End If
                                    'If locCSFlag Then
                                        If Day(xTmpEndDate) >= 15 Then
                                            xTmpDays = xTmpDays + 1
                                            'If CVDate(xTmpStartDate) = CVDate(xTmFirstDay) Then
                                            '    Debug.Print xEmpnbr
                                            'End If
                                        End If
                                    'End If
                                End If
                            End If
                            '==============None 15DaysLogic End ====
                            
                        End If
                        
                        If xTmpDays <= 0 Then xTmpDays = 0
                        If xTmpDays > xTmpMaxMth Then xTmpDays = xTmpMaxMth '12
                        
                        'check xExpYears (#/years)
                        If xExpYears > 0 Then
                            'If the employee's experience years is great than xExpYears
                            'and then give the full credited service
                            xEmpExpYears = GetEmpExpYears(xEmpnbr, xTermID)
                            If xEmpExpYears >= xExpYears Then
                                If xTmpDays < xTmpMaxMth Then
                                    xTmpDays = xTmpMaxMth '12
                                End If
                            End If
                        End If
                    Else
                        xTmpDays = 0
                    End If
                End If
            'Else 'No date found, give 0 month
            '    'xRetVal = 0
            End If
            If xTmpDays < 0 Then xTmpDays = 0
            xRetVal = xRetVal + xTmpDays
        End If
Next_Rec:
        rsEWrk.MoveNext
    Loop
    rsEWrk.Close
    
    If xRetVal > 12 Then
        xRetVal = 12
    End If
    
    'If this employee was hire less then 90 days (3 months), no CS
    If DateDiff("D", xDOH, xLastDay) < 90 Then
        xRetVal = 0
    End If
    
end_line:
    CalcCreditService = xRetVal
End Function

Public Function GetStatusDate(xEmpnbr, xCode, xDATE, xNew, Optional xTermID)
Dim rsETemp As New ADODB.Recordset
Dim SQLQ As String
Dim xRetVal
    xRetVal = ""
    If IsMissing(xTermID) Then 'active
        SQLQ = "SELECT * FROM HREMPHIS WHERE EE_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND NOT (EE_OLDSTAT IS NULL) "
        If xNew = "Y" Then
            SQLQ = SQLQ & "AND EE_NEWSTAT = '" & xCode & "' "
        End If
        If xNew = "N" Then
            SQLQ = SQLQ & "AND EE_OLDSTAT = '" & xCode & "' "
        End If
        If xRule = "CSRULES" Then
            SQLQ = SQLQ & "AND EE_OLDSTAT in (SELECT DISTINCT PE_EMPSTATUS FROM HRP_CREDITED_SERVICE_RULES WHERE PE_BREAK_IN_SERVICE = 'N') "
        End If
        SQLQ = SQLQ & "AND EE_CHGDATE <= " & Date_SQL(xDATE) & " "
        SQLQ = SQLQ & "ORDER BY EE_CHGDATE DESC "
        rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsETemp.EOF Then
            xRetVal = rsETemp("EE_CHGDATE")
        End If
        rsETemp.Close
    Else 'Term
        SQLQ = "SELECT * FROM Term_HREMPHIS WHERE EE_EMPNBR = " & xEmpnbr & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTermID & " "
        SQLQ = SQLQ & "AND NOT (EE_OLDSTAT IS NULL) "
        If xNew = "Y" Then
            SQLQ = SQLQ & "AND EE_NEWSTAT = '" & xCode & "' "
        End If
        If xNew = "N" Then
            SQLQ = SQLQ & "AND EE_OLDSTAT = '" & xCode & "' "
        End If
        If xRule = "CSRULES" Then
            SQLQ = SQLQ & "AND EE_OLDSTAT in (SELECT DISTINCT PE_EMPSTATUS FROM HRP_CREDITED_SERVICE_RULES WHERE PE_BREAK_IN_SERVICE = 'N') "
        End If
        SQLQ = SQLQ & "AND EE_CHGDATE <= " & Date_SQL(xDATE) & " "
        SQLQ = SQLQ & "ORDER BY EE_CHGDATE "
        rsETemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsETemp.EOF Then
            xRetVal = rsETemp("EE_CHGDATE")
        End If
        rsETemp.Close
    
    End If
    GetStatusDate = xRetVal
End Function

Public Function GetEmpExpYears(xEmpNo, Optional xTermID)
'Experience Years should be: from Date of Hire to the Last Date of status codes
'which the PE_BREAK_IN_SERVICE is 'N', such as ACT, FMLA,MAT
'e.g. Employee status changed from ACT to STD (12 month), and then to LTD
'SELECT DISTINCT PE_EMPSTATUS FROM HRP_CREDITED_SERVICE_RULES WHERE PE_BREAK_IN_SERVICE = 'N'
Dim rsETemp As New ADODB.Recordset
Dim SQLQ As String
Dim xTmDate
Dim xRetVal
    xRetVal = 0
    xTmDate = GetStatusDate(xEmpNo, "", xLastDay, "CSRULES", xTermID)
    If IsDate(xTmDate) Then
        xRetVal = DateDiff("d", xDOH, xTmDate)
        xRetVal = Round((xRetVal / 365), 2)
    End If
    GetEmpExpYears = xRetVal
End Function

Public Sub Upt_PayrollTransaction(xEmpNo, xINDICATOR, xCode, xStartDate, xEndDate, xAmt)
Dim rsPayTran As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HR_PAYROLL_TRANSACTION WHERE PT_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND PT_INDICATOR = '" & xINDICATOR & "' "
    SQLQ = SQLQ & "AND PT_PAYCODE = '" & xCode & "' "
    SQLQ = SQLQ & "AND PT_PAYSTART = " & Date_SQL(xStartDate) & " "
    rsPayTran.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsPayTran.EOF Then
        rsPayTran.AddNew
        rsPayTran("PT_EMPNBR") = xEmpNo
        rsPayTran("PT_INDICATOR") = xINDICATOR
        rsPayTran("PT_PAYCODE") = xCode
    End If
    rsPayTran("PT_DOLLARAMT") = xAmt
    rsPayTran("PT_PAYSTART") = xStartDate
    rsPayTran("PT_PAYEND") = xEndDate
    rsPayTran("PT_LUSER") = glbUserID
    rsPayTran("PT_LDATE") = Date
    rsPayTran("PT_LTIME") = Time$
    rsPayTran.Update
    rsPayTran.Close
End Sub
Public Sub Upt_WFCPENSIOND_NAME(xEmpNo, xSIN, xEligible, xOldVal, xNewVal, xType)
Dim SQLQ As String
Dim xField As String
    If IsNull(xEligible) Then
        Exit Sub
    End If
    If Not xEligible = "Y" Then
        Exit Sub
    End If
    xField = ""
    If xType = "Surname" Then
        xField = "PE_SURNAME"
    End If
    If xType = "Fname" Then
        xField = "PE_FNAME"
    End If
    If Len(xField) = 0 Then Exit Sub
    
    SQLQ = "UPDATE HRP_PA_DETAILS SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PA_MASTER SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PA_MASTER_ARC SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PAR SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PAR_ARC SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_ALERTS SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_BENEFICIARY SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_BENEFICIARY_ARC SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_MASTER SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_MASTER_ARC SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    'Ticket #19988 Franks 06/06/2011
    SQLQ = "UPDATE HRP_SUNLIFE SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_SUNLIFE_ARC SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    'Ticket #19988 Franks 09/06/2011
    SQLQ = "UPDATE HRP_PA_DETAILS_ARC SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_MASTER_AUDIT SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_SPECIAL_EARLY_RET SET " & xField & " = '" & xNewVal & "' WHERE PE_SIN = '" & xSIN & "' "
    gdbAdoIhr001.Execute SQLQ
End Sub

Public Sub Upt_WFCPENSIOND_SIN(xEmpNo, xEligible, xOldSIN, xNewSIN)
Dim SQLQ As String
    If IsNull(xEligible) Then
        Exit Sub
    End If
    If Not xEligible = "Y" Then
        Exit Sub
    End If
    If Len(xOldSIN) = 0 Then
        Exit Sub
    End If
    If Len(xNewSIN) = 0 Then
        Exit Sub
    End If
    SQLQ = "UPDATE HRP_PA_DETAILS SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PA_MASTER SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PA_MASTER_ARC SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PAR SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PAR_ARC SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_ALERTS SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_BENEFICIARY SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_BENEFICIARY_ARC SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_MASTER SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HRP_PENSION_MASTER_ARC SET PE_SIN = '" & xNewSIN & "' WHERE PE_SIN = '" & xOldSIN & "' "
    gdbAdoIhr001.Execute SQLQ

End Sub
Public Function get_EmpOtherByField(xEmpNo, xFieldName, Optional xTERM_Seq)
Dim rsEmOther2 As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = Null
    If IsMissing(xTERM_Seq) Then
        SQLQ = "SELECT " & xFieldName & " FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT " & xFieldName & " FROM Term_HREMP_OTHER WHERE TERM_SEQ = " & xTERM_Seq & " "
    End If
    rsEmOther2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmOther2.EOF Then
        retVal = rsEmOther2(xFieldName)
    End If
    rsEmOther2.Close
    get_EmpOtherByField = retVal
End Function
Public Sub Upt_EmpOtherByField(xEmpNo, xFieldName, xVal, Optional uptNull = "N")
Dim rsEmOther2 As New ADODB.Recordset
Dim SQLQ As String
        SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & " "
        rsEmOther2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If uptNull = "Y" Then
            If Not rsEmOther2.EOF Then
                If IsNull(rsEmOther2(xFieldName)) Then
                    rsEmOther2(xFieldName) = xVal
                    rsEmOther2.Update
                    rsEmOther2.Close
                End If
            End If
        Else
            If rsEmOther2.EOF Then
                rsEmOther2.AddNew
                rsEmOther2("ER_EMPNBR") = xEmpNo
            End If
            rsEmOther2(xFieldName) = xVal
            rsEmOther2.Update
            rsEmOther2.Close
        End If
End Sub
Public Sub Upt_PENSIONDATE2(xEmpNo, xAction, Optional xDATE)
Dim rsEmOther2 As New ADODB.Recordset
Dim SQLQ As String
    If xAction = "UPDATE" Then
        SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & " "
        rsEmOther2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsEmOther2.EOF Then
            rsEmOther2.AddNew
            rsEmOther2("ER_EMPNBR") = xEmpNo
        End If
        If Not IsMissing(xDATE) Then
            rsEmOther2("ER_PENSIONDATE2") = xDATE 'dlpDate(15)
        End If
        rsEmOther2.Update
        rsEmOther2.Close
    End If
    If xAction = "DELETE" Then
        SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & " "
        rsEmOther2.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmOther2.EOF Then
            If Not IsNull(rsEmOther2("ER_PENSIONDATE2")) Then
                rsEmOther2("ER_PENSIONDATE2") = Null
                rsEmOther2.Update
            End If
        End If
        rsEmOther2.Close
    End If
End Sub

Public Function isEmpLOA(xEmpNo)
Dim rsLEmp As New ADODB.Recordset
Dim rsLTABL As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT ED_EMPNBR,ED_EMP FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsLEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLEmp.EOF Then
        If Not IsNull(rsLEmp("ED_EMP")) Then
            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'EDEM' "
            SQLQ = SQLQ & "AND TB_KEY = '" & rsLEmp("ED_EMP") & "' "
            rsLTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsLTABL.EOF Then
                If Not rsLTABL("TB_USR3") = 0 Then
                    retVal = True
                End If
            End If
            rsLTABL.Close
        End If
    End If
    isEmpLOA = retVal
End Function

Public Function calDBPA(xDeemEarns, Optional xSalHrl)
Dim xReturn
Dim xHourly As String
    xReturn = 0
    xHourly = ""
    If Not IsMissing(xSalHrl) Then
        If xSalHrl = "Hourly" Then
            xHourly = "Y"
        End If
    End If
    If IsNumeric(xDeemEarns) Then
        If xHourly = "Y" Then
            xReturn = (xDeemEarns * 9) - 600
        Else
            xReturn = (0.009 * xDeemEarns * 9) - 600
        End If
    End If
    If xReturn < 0 Then xReturn = 0
    xReturn = Round(xReturn, 2)
    calDBPA = xReturn
End Function

Public Function GetCountryFromDiv(xDivCode)
Dim rsTTemp As New ADODB.Recordset
Dim SQLQ As String
Dim xRetVal As String
    xRetVal = ""
    SQLQ = "SELECT DV_COUNTRY FROM HR_DIVISION WHERE DIV = '" & xDivCode & "' "
    rsTTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTTemp.EOF Then
        If Not IsNull(rsTTemp("DV_COUNTRY")) Then xRetVal = rsTTemp("DV_COUNTRY")
    End If
    rsTTemp.Close
    GetCountryFromDiv = xRetVal
End Function

Public Sub glbUpdateBenefitGroup(xEmpNo, SaveBGroup, NewBGroup, xEDate)
Dim rsBGMST As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim rsBGEE As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String
Dim BelongOldGroup As Boolean
gdbAdoIhr001W.BeginTrans
gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
gdbAdoIhr001W.CommitTrans

'Len(NewBGroup) = 0, it means deleting the Benefit Group
If Len(NewBGroup) > 0 Then
    gdbAdoIhr001W.BeginTrans
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & NewBGroup & "' "
    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsBGMST.EOF
        rsBGTMP.AddNew
        rsBGTMP("BM_COMPNO") = "001"
        rsBGTMP("BM_BENEFIT_GROUP") = NewBGroup
        rsBGTMP("BM_BCODE") = rsBGMST("BM_BCODE")
        If glbWFC Then
            rsBGTMP("BM_EDATE") = xEDate
            If rsBGMST("BM_BCODE") = "LIF1" Or rsBGMST("BM_BCODE") = "LIF3" Then
                rsBGTMP("BM_CHECK") = 0
            Else
            rsBGTMP("BM_CHECK") = 1
            End If
        Else
            rsBGTMP("BM_EDATE") = rsBGMST("BM_EDATE")
            rsBGTMP("BM_CHECK") = 1
        End If
        rsBGTMP("BM_COVER") = rsBGMST("BM_COVER")
        rsBGTMP("BM_AMT") = rsBGMST("BM_AMT")
        rsBGTMP("BM_PPAMT") = rsBGMST("BM_PPAMT")
        rsBGTMP("BM_UNITCOST") = rsBGMST("BM_UNITCOST")
        rsBGTMP("BM_PCE") = rsBGMST("BM_PCE")
        rsBGTMP("BM_PCC") = rsBGMST("BM_PCC")
        rsBGTMP("BM_ECOST") = rsBGMST("BM_ECOST")
        rsBGTMP("BM_CCOST") = rsBGMST("BM_CCOST")
        rsBGTMP("BM_TCOST") = rsBGMST("BM_TCOST")
        rsBGTMP("BM_MAXDOL") = rsBGMST("BM_MAXDOL")
        rsBGTMP("BM_PREMIUM") = rsBGMST("BM_PREMIUM")
        rsBGTMP("BM_PER") = rsBGMST("BM_PER")
        rsBGTMP("BM_MTHCCOST") = rsBGMST("BM_MTHCCOST")
        rsBGTMP("BM_MTHECOST") = rsBGMST("BM_MTHECOST")
        rsBGTMP("BM_TAXBEN") = rsBGMST("BM_TAXBEN")
        rsBGTMP("BM_SALARYDEPENDANT") = rsBGMST("BM_SALARYDEPENDANT")
        rsBGTMP("BM_MINIMUM") = rsBGMST("BM_MINIMUM")
        rsBGTMP("BM_FACTOR") = rsBGMST("BM_FACTOR")
        rsBGTMP("BM_ROUND") = rsBGMST("BM_ROUND")
        rsBGTMP("BM_MAXIMUM") = rsBGMST("BM_MAXIMUM")
        rsBGTMP("BM_NEXTNEAREST") = rsBGMST("BM_NEXTNEAREST")
        rsBGTMP("BM_TAXAMOUNT") = rsBGMST("BM_TAXAMOUNT")
        rsBGTMP("BM_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
        
        rsBGTMP("BM_DWM") = rsBGMST("BM_DWM")
        rsBGTMP("BM_PERORDOLL") = rsBGMST("BM_PERORDOLL")
        
        rsBGTMP("BM_POLICY") = rsBGMST("BM_POLICY")
        
        rsBGTMP("BM_COMMENTS") = rsBGMST("BM_COMMENTS")
        rsBGTMP("BM_PTAX") = rsBGMST("BM_PTAX")
        'rsBGTMP("BM_CHECK") = 1
        rsBGTMP("BM_ACTION") = "Add"
        rsBGTMP("BM_WRKEMP") = glbUserID
        
        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BM_BCODE") & "' "
        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTABL.EOF Then
            rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
        End If
        rsTABL.Close
        rsBGTMP.Update
        rsBGMST.MoveNext
    Loop
    rsBGTMP.Close
    rsBGMST.Close
    gdbAdoIhr001W.CommitTrans
    If Not glbSQL And Not glbOracle Then Call Pause(1)
    
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo
    
    rsBGEE.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    gdbAdoIhr001W.BeginTrans
    Do Until rsBGEE.EOF
        SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID
        SQLQ = SQLQ & "' AND  BM_BCODE='" & rsBGEE("BF_BCODE") & "'"
        SQLQ = SQLQ & " AND BM_ACTION='Add' " 'Frank 11/04/2003 for Duplicate record entere and delete
        rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenStatic, adLockOptimistic
        If rsBGTMP.EOF Then
            BelongOldGroup = False
            If rsBGEE("BF_GROUP") = SaveBGroup Then
                BelongOldGroup = True
            Else
                SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & SaveBGroup & "'"
                SQLQ = SQLQ & " AND  BM_BCODE='" & rsBGEE("BF_BCODE")
                If IsNull(rsBGEE("BF_COVER")) Then
                    SQLQ = SQLQ & "' AND (BM_COVER IS NULL OR BM_COVER='')"
                Else
                    SQLQ = SQLQ & "' AND BM_COVER='" & rsBGEE("BF_COVER") & "'"
                End If
                rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsBGMST.EOF Then
                    BelongOldGroup = True
                End If
                rsBGMST.Close
            End If
            If BelongOldGroup Then
                rsBGTMP.AddNew
                rsBGTMP("BM_BCODE") = rsBGEE("BF_BCODE")
                rsBGTMP("BM_COVER") = rsBGEE("BF_COVER")
                rsBGTMP("BM_EDATE") = rsBGEE("BF_EDATE")
                rsBGTMP("BM_BENEFIT_GROUP") = SaveBGroup
                rsBGTMP("BM_CHECK") = 1
                If glbWFC Then 'Ticket #18810
                    rsBGTMP("BM_ACTION") = "EndDate"
                    rsBGTMP("BM_BENEFIT_GROUP") = Null 'delete old Benefit Group
                    rsBGTMP("BM_POLICY") = rsBGEE("BF_POLICY")
                Else
                    rsBGTMP("BM_ACTION") = "Delete"
                End If
                rsBGTMP("BM_WRKEMP") = glbUserID
                rsBGTMP("BM_WRKID") = rsBGEE("BF_BENE_ID")
                SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGEE("BF_BCODE") & "' "
                rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
                If Not rsTABL.EOF Then
                    rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
                End If
                rsTABL.Close
                rsBGTMP("BM_CHECK") = 1
                rsBGTMP.Update
            End If
            
        Else
            rsBGTMP("BM_WRKID") = rsBGEE("BF_BENE_ID")
            rsBGTMP("BM_CHECK") = 1
            rsBGTMP("BM_ACTION") = "Update"
            rsBGTMP.Update
        End If
        rsBGTMP.Close
        rsBGEE.MoveNext
    Loop
    gdbAdoIhr001W.CommitTrans
Else 'Deleting the Benefit Group
    gdbAdoIhr001W.BeginTrans
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND BF_GROUP ='" & SaveBGroup & "' "
    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsBGMST.EOF
        rsBGTMP.AddNew
        rsBGTMP("BM_BCODE") = rsBGMST("BF_BCODE")
        rsBGTMP("BM_COVER") = rsBGMST("BF_COVER")
        rsBGTMP("BM_EDATE") = rsBGMST("BF_EDATE")
        rsBGTMP("BM_BENEFIT_GROUP") = SaveBGroup
        rsBGTMP("BM_CHECK") = 1
        If glbWFC Then 'Ticket #18810
            rsBGTMP("BM_ACTION") = "EndDate"
            rsBGTMP("BM_POLICY") = rsBGMST("BF_POLICY")
        Else
            rsBGTMP("BM_ACTION") = "Delete"
        End If
        rsBGTMP("BM_WRKEMP") = glbUserID
        rsBGTMP("BM_WRKID") = rsBGMST("BF_BENE_ID")
        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BF_BCODE") & "' "
        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTABL.EOF Then
            rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
        End If
        rsTABL.Close
        rsBGTMP("BM_CHECK") = 1
        rsBGTMP.Update
        rsBGMST.MoveNext
    Loop
    rsBGTMP.Close
    rsBGMST.Close
    gdbAdoIhr001W.CommitTrans
End If

End Sub

Public Function WFC_AUDIT_MANULIFE_BENF(xEmpNo, xBCode, xBEDate, xBCover, xPolicy, xBenEndDate) 'No AU_CEASEDATE in HRAUDIT, Jerry said we will add it in next release
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
'''On Error GoTo AUDIT_ERR
WFC_AUDIT_MANULIFE_BENF = False

If IsNull(xPolicy) Then
    Exit Function
End If
If Len(xPolicy) = "" Then
    Exit Function
End If

'BENEFIT End Date
If IsDate(xBenEndDate) Then
    If OBenEndDate = "" Then
        GoTo MODUPD
    Else
        If IsDate(OBenEndDate) Then
            If CVDate(xBenEndDate) <> CVDate(OBenEndDate) Then 'Ticket #15591
                GoTo MODUPD
            End If
        End If
    End If
Else
    If IsDate(OBenEndDate) Then
        GoTo MODUPD
    End If
End If

GoTo MODNOUPD

MODUPD:

rsTB.Open "SELECT ED_DIV, ED_SECTION, ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1  FROM HREMP WHERE ED_EMPNBR = " & xEmpNo, gdbAdoIhr001, adOpenKeyset
If rsTB.EOF Then
    rsTB.Close:    GoTo MODNOUPD
End If
If IsNull(rsTB("ED_USER_TEXT1")) Then 'Certificate #
    rsTB.Close:    GoTo MODNOUPD
Else
    If Len(Trim(rsTB("ED_USER_TEXT1"))) = 0 Then
        rsTB.Close:    GoTo MODNOUPD
    End If
End If

rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

rsTA.AddNew
rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
rsTA("MT_PT_TABL") = "EDPT"
rsTA("MT_TYPE") = "T"
rsTA("MT_BENEFIT") = xBCode
rsTA("MT_EDATE") = xBEDate
If IsDate(xBenEndDate) Then
rsTA("MT_CEASEDATE") = xBenEndDate
End If
If Len(xBCover) > 0 Then rsTA("MT_COVER") = xBCover
If Len(Trim(xPolicy)) > 0 Then
    rsTA("MT_POLICY_NO") = Trim(xPolicy)
End If
rsTA("MT_COMPNO") = "001"
rsTA("MT_EMPNBR") = xEmpNo
rsTA("MT_ACCOUNT_NO") = rsTB("ED_USER_NUM1")
rsTA("MT_CERT_NO") = rsTB("ED_USER_TEXT1")
rsTA("MT_COVERAGE_CLASS") = rsTB("ED_USER_TEXT2")
rsTA("MT_UPLOAD") = "N"
rsTA("MT_LUSER") = glbUserID
If Not IsDate(xBenEndDate) Then
    rsTA("MT_LDATE") = Format(Date, "SHORT DATE")
Else
    If CVDate(xBenEndDate) < CVDate(Date) Then 'WFC Ticket #14867
        rsTA("MT_LDATE") = Format(Date, "SHORT DATE")
    Else
        rsTA("MT_LDATE") = Format(xBenEndDate, "SHORT DATE")
    End If
End If
rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
rsTA("MT_LTIME") = Time$

rsTA.Update

MODNOUPD:
WFC_AUDIT_MANULIFE_BENF = True
Exit Function
AUDIT_ERR:

'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING MANULIFE AUDIT RECORD", "MANULIFE AUDIT FILE", "UPDATE")
'If gintRollBack% = False Then Resume Next Else Unload Me
Resume Next

End Function

Public Function getNGSSubGrpFromMatrix(xDIV, xUnion)
Dim rsNGS As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    If glbNGS_OnFlag Then
        If Len(xDIV) > 0 And Len(xUnion) > 0 Then
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
            SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
            rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsNGS.EOF Then
                retVal = rsNGS("NG_SUB_GROUP")
            End If
        End If
    End If
    getNGSSubGrpFromMatrix = retVal
End Function
Public Function getNGSPayGrpFromMatrix(xDIV, xUnion)
Dim rsNGS As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    If glbNGS_OnFlag Then
        If Len(xDIV) > 0 And Len(xUnion) > 0 Then
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
            SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
            rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsNGS.EOF Then
                retVal = rsNGS("NG_PAY_GROUP")
            End If
        End If
    End If
    getNGSPayGrpFromMatrix = retVal
End Function

'Ticket #20885 Franks 12/01/2011
Public Sub SamuelAuditAdd(xEmpNo, xType, xForm, xItem, xOldVal, xNewVal, Optional xLDate, Optional xTermSEQ)
Dim rsAuditSAM As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HRAUDIT_SAMUEL WHERE 1=2"
    rsAuditSAM.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsAuditSAM.AddNew
    rsAuditSAM("NG_EMPNBR") = xEmpNo
    rsAuditSAM("NG_TYPE") = xType
    rsAuditSAM("NG_FORM") = xForm
    rsAuditSAM("NG_ITEM") = xItem
    rsAuditSAM("NG_NEW_VALUE") = xNewVal
    rsAuditSAM("NG_OLD_VALUE") = xOldVal
    If Len(glbEmpDiv) > 0 Then rsAuditSAM("NG_DIV") = glbEmpDiv
    'If Len(glbUNION) > 0 Then rsAuditSAM("NG_ORG") = glbUNION
    If Len(glbEmpAdminBy) > 0 Then rsAuditSAM("NG_ADMINBY") = glbEmpAdminBy
    If Len(glbEmpSection) > 0 Then rsAuditSAM("NG_SECTION") = glbEmpSection
    If Len(glbEmpRegion) > 0 Then rsAuditSAM("NG_REGION") = glbEmpRegion
    rsAuditSAM("NG_UPLOAD") = "N"
    rsAuditSAM("NG_LUSER") = glbUserID
    If IsMissing(xLDate) Then
        rsAuditSAM("NG_LDATE") = Date
    Else
        rsAuditSAM("NG_LDATE") = xLDate
    End If
    rsAuditSAM("NG_LTIME") = Time$
    If IsMissing(xTermSEQ) Then
        rsAuditSAM("TERM_SEQ") = 0
    Else
        rsAuditSAM("TERM_SEQ") = xTermSEQ
    End If
    rsAuditSAM.Update
    rsAuditSAM.Close
End Sub

'Ticket #19266
Public Sub NGSAuditAdd(xEmpNo, xType, xForm, xItem, xOldVal, xNewVal, Optional xLDate, Optional xTermSEQ)
Dim rsNGSAUDIT As New ADODB.Recordset
Dim SQLQ As String
    'WFC Issues - Oct. 18-12 - Ticket #22699
    'NGS:  If an employee is going from FT to something else, the NGS Audit Record needs to be updated with the NGS End Date. If there is a NGS End Date for that record, remove it.
    If xItem = lStr("Other Date 2") Then  'NGS End Date, remove the previous record with End Date
        If IsDate(xNewVal) Then
            SQLQ = "DELETE FROM WFC_NGS_AUDIT WHERE NG_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND NG_ITEM = '" & xItem & "' "
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    
    SQLQ = "SELECT * FROM WFC_NGS_AUDIT WHERE 1=2"
    rsNGSAUDIT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsNGSAUDIT.AddNew
    rsNGSAUDIT("NG_EMPNBR") = xEmpNo
    rsNGSAUDIT("NG_TYPE") = xType
    rsNGSAUDIT("NG_FORM") = xForm
    rsNGSAUDIT("NG_ITEM") = xItem
    rsNGSAUDIT("NG_NEW_VALUE") = xNewVal
    rsNGSAUDIT("NG_OLD_VALUE") = xOldVal
    If Len(glbEmpDiv) > 0 Then rsNGSAUDIT("NG_DIV") = glbEmpDiv
    If Len(glbUNION) > 0 Then rsNGSAUDIT("NG_ORG") = glbUNION
    If Len(glbWFCNGSSubGroup) > 0 Then rsNGSAUDIT("NG_SUB_GROUP") = glbWFCNGSSubGroup
    If Len(glbWFCPayGroup) > 0 Then rsNGSAUDIT("NG_PAY_GROUP") = glbWFCPayGroup
    rsNGSAUDIT("NG_UPLOAD") = "N"
    rsNGSAUDIT("NG_LUSER") = glbUserID
    If IsMissing(xLDate) Then
        rsNGSAUDIT("NG_LDATE") = Date
    Else
        rsNGSAUDIT("NG_LDATE") = xLDate
    End If
    rsNGSAUDIT("NG_LTIME") = Time$
    If IsMissing(xTermSEQ) Then
        rsNGSAUDIT("TERM_SEQ") = 0
    Else
        rsNGSAUDIT("TERM_SEQ") = xTermSEQ
    End If
    rsNGSAUDIT.Update
    rsNGSAUDIT.Close
    'to create audit report
    'create NGS Audit view using NG_EMPNBR + TermSEQ as key
    'create HREMP UNION TERM_HREMP and use ED_EMPNBR + 0 as key for Active, ED_EMPNBR + TermSEQ as key for term
    'then this report can show employee name for both active and term

End Sub

Public Sub AUDIT_NGS_NEWHIRE(xEmpNo, Optional xLDate)
Dim rsEmp As New ADODB.Recordset
Dim rsEOther As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
Dim SQLQ As String
Dim xUptFlag As Boolean

On Error GoTo AUDIT_ERR

If Not glbNGS_OnFlag Then
    Exit Sub
End If

SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsEmp.EOF Then
    rsEmp.Close
    Exit Sub
End If

If IsMissing(xLDate) Then
    xLDate = Date
End If

'No NGS Sub Group, skip
If IsNull(rsEmp("ED_VADIM1")) Then Exit Sub
If Len(rsEmp("ED_VADIM1")) = 0 Then Exit Sub

'No NGS Effective Date, skip
SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo
rsEOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsEOther.EOF Then
    rsEOther.Close
    Exit Sub
End If
'If IsNull(rsEOther("ER_OTHERDATE1")) Then Exit Sub

'NGS field changes
'Demo fields ---------------------------
Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Surname", "", rsEmp("ED_SURNAME"), xLDate)
Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "First Name", "", rsEmp("ED_SURNAME"), xLDate)
Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "SSN", "", rsEmp("ED_SIN"), xLDate)
Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Birth Date", "", rsEmp("ED_DOB"), xLDate)
If Not IsNull(rsEmp("ED_SEX")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Gender", "", rsEmp("ED_SEX"), xLDate)
If Not IsNull(rsEmp("ED_SMOKER")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Smoker", "", IIf(rsEmp("ED_SMOKER"), "Y", "N"), xLDate)
If Not IsNull(rsEmp("ED_MSTAT")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Marital Status", "", rsEmp("ED_MSTAT"), xLDate)
If Not IsNull(rsEmp("ED_ADDR1")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Address 1", "", rsEmp("ED_ADDR1"), xLDate)
If Not IsNull(rsEmp("ED_ADDR2")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Address 2", "", rsEmp("ED_ADDR2"), xLDate)
If Not IsNull(rsEmp("ED_CITY")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "City", "", rsEmp("ED_CITY"), xLDate)
If Not IsNull(rsEmp("ED_PROV")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Province", "", rsEmp("ED_PROV"), xLDate)
If Not IsNull(rsEmp("ED_COUNTRY")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Country", "", rsEmp("ED_COUNTRY"), xLDate)
If Not IsNull(rsEmp("ED_PCODE")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Zip Code", "", rsEmp("ED_PCODE"), xLDate)
If Not IsNull(rsEmp("ED_PHONE")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Demographics", "Telephone", "", rsEmp("ED_PHONE"), xLDate)

'Status/Dates fields -------------------
'"NGS Sub Group"
Call NGSAuditAdd(glbLEE_ID, "A", "Status/Dates", lStr("Vadim Field 1"), "", rsEmp("ED_VADIM1"), xLDate)
'"Pay Group"
If Not IsNull(rsEmp("ED_VADIM2")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Status/Dates", lStr("Vadim Field 2"), "", rsEmp("ED_VADIM2"), xLDate)
Call NGSAuditAdd(glbLEE_ID, "A", "Status/Dates", lStr("Category"), "", rsEmp("ED_PT"), xLDate)
Call NGSAuditAdd(glbLEE_ID, "A", "Status/Dates", "Date of Hire", "", rsEmp("ED_DOH"), xLDate)
If Not IsNull(rsEOther("ER_OTHERDATE1")) Then Call NGSAuditAdd(glbLEE_ID, "A", "Status/Dates", lStr("Other Date 1"), "", rsEOther("ER_OTHERDATE1"), xLDate)

rsEmp.Close
rsEOther.Close

Exit Sub
AUDIT_ERR:

glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING NGS AUDIT RECORD", "NGS AUDIT FILE", "New Hire")
'If gintRollBack% = False Then Resume Next Else Unload Me

End Sub


Sub WFCTermEmailInfo(xUserID)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ '
    SQLQ = "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(xUserID, "'", "''") & "' "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("EM_SERVER")) Then
            If Len(Trim((rsTemp("EM_SERVER")))) > 0 Then
                glbSMTPServerIP = Trim((rsTemp("EM_SERVER")))
            End If
        End If
    End If
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("EM_ADDRESS")) Then
            If Len(Trim((rsTemp("EM_ADDRESS")))) > 0 Then
                glbWFCTermEmail = Trim((rsTemp("EM_ADDRESS")))
            End If
        End If
    End If
    rsTemp.Close

End Sub

Sub SetWFCSerialNo()
Dim xPlant(22, 2)
Dim xPlantCode, I
    xPlant(1, 0) = "ADD"
    xPlant(1, 1) = "ADDISON"
    xPlant(1, 2) = "S/N - 2310W"
    xPlant(2, 0) = "ATLA"
    xPlant(2, 1) = "ATLANA"
    xPlant(2, 2) = "S/N - 2305W"
    xPlant(3, 0) = "BLEN"
    xPlant(3, 1) = "BLENHEIM"
    xPlant(3, 2) = ""
    xPlant(4, 0) = "BROD"
    xPlant(4, 1) = "BRODHEAD"
    xPlant(4, 2) = "S/N - 2311W"
    xPlant(5, 0) = "CHAT"
    xPlant(5, 1) = "CHATTANOOGA"
    xPlant(5, 2) = "S/N - 2312W"
    xPlant(6, 0) = "DELR" '"DEL" '
    xPlant(6, 1) = "DEL RIO"
    xPlant(6, 2) = "S/N - 2320W"
    xPlant(7, 0) = "ELPA" '"ELP" '
    xPlant(7, 1) = "EL PASO"
    xPlant(7, 2) = "S/N - 2320W"
    xPlant(8, 0) = "FAIR"
    xPlant(8, 1) = "CARTEX -  FAIRLESS HILLS"
    xPlant(8, 2) = "S/N - 2264"
    xPlant(9, 0) = "FREM"
    xPlant(9, 1) = "FREMONT"
    xPlant(9, 2) = "S/N - 2313W"
    xPlant(10, 0) = "KC"
    xPlant(10, 1) = "KANSAS CITY"
    xPlant(10, 2) = "S/N - 2314W"
    xPlant(11, 0) = "KIPL"
    xPlant(11, 1) = "KIPLING"
    xPlant(11, 2) = "S/N - 2361W"
    xPlant(12, 0) = "MISS"
    xPlant(12, 1) = "MISSISSAUGA"
    xPlant(12, 2) = "S/N - 2382W"
    xPlant(13, 0) = "MORV"
    xPlant(13, 1) = "MORVAL"
    xPlant(13, 2) = "S/N - 2315W"
    xPlant(14, 0) = "ROM"
    xPlant(14, 1) = "ROMULUS"
    xPlant(14, 2) = "S/N - 2268W"
    xPlant(15, 0) = "SARN"
    xPlant(15, 1) = "SARNIA"
    xPlant(15, 2) = "S/N - 2286W"
    xPlant(16, 0) = "STPE"
    xPlant(16, 1) = "ST. PETERS"
    xPlant(16, 2) = "S/N - 2316W"
    xPlant(17, 0) = "TILB"
    xPlant(17, 1) = "TILBURY"
    xPlant(17, 2) = "S/N - 2317W"
    xPlant(18, 0) = "TROY"
    xPlant(18, 1) = "TROY"
    xPlant(18, 2) = "S/N - 2283W"
    xPlant(19, 0) = "WHBY"
    xPlant(19, 1) = "WHITBY"
    xPlant(19, 2) = "S/N - 2271W"
    xPlant(20, 0) = "WHLK"
    xPlant(20, 1) = "WHITMORE LAKE"
    xPlant(20, 2) = "S/N - 2281W"
    xPlant(21, 0) = "GREN" '"GRN"
    xPlant(21, 1) = "GREENSBORO"
    xPlant(21, 2) = "S/N - 2501WFC" 'Serial# Made by Frank
    xPlant(22, 0) = "EPLM" '"ELL" '
    xPlant(22, 1) = "EL PASO LAMINATION"
    xPlant(22, 2) = "S/N - 2320Z"
    
    glbPlantCode = Replace(Trim(Mid(glbSeleSection, 12, 4)), "'", "")
    For I = 1 To 22
        If glbPlantCode = xPlant(I, 0) Then
            glbPlantDesc = xPlant(I, 1)
            'If Len(xPlant(I, 2)) > 0 Then
            '    glbCompSerial = xPlant(I, 2)
            'End If
        End If
    Next I
    
End Sub

Public Function getUserSecList()
Dim xList
Dim I As Integer
Dim K As Integer
Dim retVal
    retVal = ""
    xList = glbSeleSection
    If InStr(1, glbSeleSection, "TB_KEY=") > 0 Then
        Do While InStr(1, xList, "TB_KEY=") > 0
            I = InStr(1, xList, "TB_KEY=")
            K = Len(xList)
            If Len(retVal) = 0 Then
                retVal = Trim(Mid(xList, I + 7, 6))
            Else
                retVal = retVal & "," & Trim(Mid(xList, I + 7, 6))
            End If
            xList = Trim(Mid(xList, I + 7 + 6, K))
        Loop
    End If
    If Len(retVal) > 0 Then
        retVal = "(" & retVal & ")"
    End If
    getUserSecList = retVal
End Function

Public Function getMSDesc(xCode As String)
Dim xRetVal As String
    xRetVal = ""
    Select Case xCode
    Case "S"
        xRetVal = "Single"
    Case "M"
        xRetVal = "Married"
    Case "F"
        xRetVal = "Family"
    Case "P"
        xRetVal = "Parent(Single)"
    Case "D"
        xRetVal = "Divorced"
    Case "W"
        xRetVal = "Widowed"
    Case "C"
        xRetVal = "Common-Law"
    Case "R"
        xRetVal = "Partner"
    Case "X"
        xRetVal = "Same-Sex"
    Case "O"
        xRetVal = "Other"
    Case "A"
        xRetVal = "Separated"
    End Select
    getMSDesc = xRetVal
End Function

Public Sub AUDIT_PENSION_MASTER(rsTmpPen As ADODB.Recordset)
Dim rsTA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
Dim SQLQ As String
Dim xBenTermDate
Dim xEmpStatCode
Dim xUptFlag As Boolean
Dim xDate1, xDate2
Dim xLDate
Dim xForm As String
'Dim UpdPenAudit As Boolean
Dim toEMPNBR, toSIN, toSURNAME, toFNAME, toITEM
Dim toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL
Dim nDBStatus
Dim nDBStaDate
Dim nEntryDate
Dim nExitDate
Dim nEarnPen
Dim nContSrv
Dim nCredSrv
Dim nBenRate
Dim IsDelete As Boolean

On Error GoTo AUDIT_ERR

If rsTmpPen.EOF Then
    Exit Sub
End If

toEMPNBR = rsTmpPen("PE_EMPNBR")
toSIN = rsTmpPen("PE_SIN")
toSURNAME = rsTmpPen("PE_SURNAME")
toFNAME = rsTmpPen("PE_FNAME")
toSECTION = rsTmpPen("PE_SECTION")
toYEARDATE = rsTmpPen("PE_YEAR_DATE")
toPENSIONTYPE = rsTmpPen("PE_PENSIONTYPE")
toHRLYSAL = rsTmpPen("PE_HRLYSAL")

If IsNull(rsTmpPen("PE_DB_STATUS")) Then nDBStatus = "" Else nDBStatus = rsTmpPen("PE_DB_STATUS")
If IsNull(rsTmpPen("PE_DB_STATUS_DATE")) Then nDBStaDate = "" Else nDBStaDate = rsTmpPen("PE_DB_STATUS_DATE")
If IsNull(rsTmpPen("PE_PEN_ENTRY_DATE")) Then nEntryDate = "" Else nEntryDate = rsTmpPen("PE_PEN_ENTRY_DATE")
If IsNull(rsTmpPen("PE_EXIT_DATE")) Then nExitDate = "" Else nExitDate = rsTmpPen("PE_EXIT_DATE")
If IsNull(rsTmpPen("PE_YEAR_AMOUNT")) Then nEarnPen = "" Else nEarnPen = rsTmpPen("PE_YEAR_AMOUNT")
If IsNull(rsTmpPen("PE_CONT_SERV")) Then nContSrv = "" Else nContSrv = rsTmpPen("PE_CONT_SERV")
If IsNull(rsTmpPen("PE_CREDITED_SERV")) Then nCredSrv = "" Else nCredSrv = rsTmpPen("PE_CREDITED_SERV")
If IsNull(rsTmpPen("PE_BENEFIT_RATE")) Then nBenRate = "" Else nBenRate = rsTmpPen("PE_BENEFIT_RATE")


If toTYPE = "Change" Then
    If Not oDBStatus = nDBStatus Then UpdPenAudit = True
    If Not oDBStaDate = nDBStaDate Then UpdPenAudit = True
    If Not oEntryDate = nEntryDate Then UpdPenAudit = True
    If Not oExitDate = nExitDate Then UpdPenAudit = True
    If Not oEarnPen = nEarnPen Then UpdPenAudit = True
    If Not oContSrv = nContSrv Then UpdPenAudit = True
    If Not oCredSrv = nCredSrv Then UpdPenAudit = True
    If Not oBenRate = nBenRate Then UpdPenAudit = True
End If

If Not UpdPenAudit Then
    Exit Sub      'No change
End If


If Not oDBStatus = nDBStatus Then
    toITEM = "Pension Status"
    toNEWVAL = nDBStatus
    toOLDVAL = oDBStatus
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If
'Effective Date
'If IsDate(dlpDate(0).Text) Then xDate1 = CVDate(dlpDate(0).Text) Else xDate1 = ""
'If IsDate(oDBStaDate) Then xDate2 = CVDate(oDBStaDate) Else xDate2 = ""
If Not (oDBStaDate = nDBStaDate) Then
    toITEM = "Effective Date"
    toNEWVAL = nDBStaDate: toOLDVAL = oDBStaDate
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If
'Membership Entry Date
'If IsDate(dlpDate(2).Text) Then xDate1 = CVDate(dlpDate(2).Text) Else xDate1 = ""
'If IsDate(oEntryDate) Then xDate2 = CVDate(oEntryDate) Else xDate2 = ""
If Not (oEntryDate = nEntryDate) Then
    toITEM = "Membership Entry Date"
    toNEWVAL = nEntryDate: toOLDVAL = oEntryDate
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If
'Membership Exit
'If IsDate(dlpDate(1).Text) Then xDate1 = CVDate(dlpDate(1).Text) Else xDate1 = ""
'If IsDate(oExitDate) Then xDate2 = CVDate(oExitDate) Else xDate2 = ""
If Not (oExitDate = nExitDate) Then
    toITEM = "Membership Exit Date"
    toNEWVAL = nExitDate: toOLDVAL = oExitDate
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If
'Earned Pension
If Not oEarnPen = nEarnPen Then
    toITEM = "Earned Pension"
    toNEWVAL = nEarnPen
    toOLDVAL = oEarnPen
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If
'Continuous Service
If Not oContSrv = nContSrv Then
    toITEM = "Continuous Service"
    toNEWVAL = nContSrv
    toOLDVAL = oContSrv
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If
'Credited Service
If Not oCredSrv = nCredSrv Then
    toITEM = "Credited Service"
    toNEWVAL = nCredSrv
    toOLDVAL = oCredSrv
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If
'Benefit Rate
If Not oBenRate = nBenRate Then
    toITEM = "Benefit Rate"
    toNEWVAL = nBenRate
    toOLDVAL = oBenRate
    Call PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
End If

Exit Sub
AUDIT_ERR:

'glbFrmCaption$ = Me.Caption
glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, "", "ADDING PEN AUDIT RECORD", "PEN AUDIT FILE", "UPDATE")
'If gintRollBack% = False Then Resume Next Else Unload Me
MsgBox Err.Description

End Sub

Public Sub PenMasterAuditAdd(toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM, toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE, toHRLYSAL)
Dim rsPenAudit As New ADODB.Recordset
Dim SQLQ As String
'Dim toEMPNBR, toSIN, toSURNAME, toFNAME, toTYPE, toSOURCE, toITEM
'Dim toNEWVAL, toOLDVAL, toSECTION, toPENSIONTYPE, toYEARDATE
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER_AUDIT WHERE 1=2"
    rsPenAudit.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsPenAudit.AddNew
    rsPenAudit("PE_EMPNBR") = toEMPNBR
    rsPenAudit("PE_SIN") = toSIN
    rsPenAudit("PE_SURNAME") = toSURNAME
    rsPenAudit("PE_FNAME") = toFNAME
    rsPenAudit("PE_TYPE") = toTYPE
    rsPenAudit("PE_SOURCE") = Left(toSOURCE, 30)
    rsPenAudit("PE_ITEM") = toITEM
    rsPenAudit("PE_NEW_VALUE") = toNEWVAL
    rsPenAudit("PE_OLD_VALUE") = toOLDVAL
    rsPenAudit("PE_SECTION") = toSECTION
    rsPenAudit("PE_PENSIONTYPE") = toPENSIONTYPE
    rsPenAudit("PE_YEAR_DATE") = toYEARDATE
    rsPenAudit("PE_HRLYSAL") = toHRLYSAL
    rsPenAudit("PE_LUSER") = glbUserID
    rsPenAudit("PE_LDATE") = Date
    rsPenAudit("PE_LTIME") = Time$
    rsPenAudit.Update
    rsPenAudit.Close
End Sub

Public Sub PenMasterAuditOldValSetup(xSetType, Optional locPenRS As ADODB.Recordset)
    If xSetType = "Blank" Then
        oDBStatus = ""
        oDBStaDate = ""
        oEntryDate = ""
        oExitDate = ""
        oEarnPen = ""
        oContSrv = ""
        oCredSrv = ""
        oBenRate = ""
    End If
    If xSetType = "CurValues" Then
        If IsNull(locPenRS("PE_DB_STATUS")) Then oDBStatus = "" Else oDBStatus = locPenRS("PE_DB_STATUS")
        If IsNull(locPenRS("PE_DB_STATUS_DATE")) Then oDBStaDate = "" Else oDBStaDate = locPenRS("PE_DB_STATUS_DATE")
        If IsNull(locPenRS("PE_PEN_ENTRY_DATE")) Then oEntryDate = "" Else oEntryDate = locPenRS("PE_PEN_ENTRY_DATE")
        If IsNull(locPenRS("PE_EXIT_DATE")) Then oExitDate = "" Else oExitDate = locPenRS("PE_EXIT_DATE")
        If IsNull(locPenRS("PE_YEAR_AMOUNT")) Then oEarnPen = "" Else oEarnPen = locPenRS("PE_YEAR_AMOUNT")
        If IsNull(locPenRS("PE_CONT_SERV")) Then oContSrv = "" Else oContSrv = locPenRS("PE_CONT_SERV")
        If IsNull(locPenRS("PE_CREDITED_SERV")) Then oCredSrv = "" Else oCredSrv = locPenRS("PE_CREDITED_SERV")
        If IsNull(locPenRS("PE_BENEFIT_RATE")) Then oBenRate = "" Else oBenRate = locPenRS("PE_BENEFIT_RATE")
    End If
End Sub

' // Returns the list index value for the item matching strSearchValue
' // in combo box cbComboBox
' // Otherwise returns -1 if not found
Public Function FindCBIndex(ByRef cbComboBox As ComboBox, ByRef strSearchValue As String, Optional xlen = 0) As Integer
    Dim n As Integer
    For n = 0 To cbComboBox.ListCount - 1
        If IIf(xlen = 0, cbComboBox.List(n), Left(cbComboBox.List(n), xlen)) = strSearchValue Then
          ' // Return the found index
            FindCBIndex = n
          ' // and exit
            Exit Function
        End If
    Next
  ' // Set not found value
    FindCBIndex = -1
End Function

Public Sub uptEmpDates(xEmpNo, xField, xVal)
Dim SQLQ As String
If IsDate(xVal) Then
    SQLQ = "UPDATE HREMP SET " & xField & " = " & Date_SQL(xVal) & " WHERE ED_EMPNBR = " & xEmpNo
    gdbAdoIhr001.Execute SQLQ
End If
If IsNull(xVal) Then
    SQLQ = "UPDATE HREMP SET " & xField & " = null " & " WHERE ED_EMPNBR = " & xEmpNo
    gdbAdoIhr001.Execute SQLQ
End If
End Sub

Public Function getPenStatusFromHRTABL(xCode)
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'EDEM' "
    SQLQ = SQLQ & "AND TB_KEY = '" & xCode & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTABL.EOF Then
        If Not IsNull(rsTABL("TB_USR1")) Then
            If Len(rsTABL("TB_USR1")) > 0 Then
                retVal = rsTABL("TB_USR1")
            End If
        End If
    End If
    rsTABL.Close
    getPenStatusFromHRTABL = retVal
End Function

Public Function getWFCPenStatus(xEmpNo) 'Ticket #23361 Franks 03/11/2013
Dim rsPen As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC "
    If rsPen.State <> 0 Then rsPen.Close
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsPen.EOF Then
        If Not IsNull(rsPen("PE_DB_STATUS")) Then
            retVal = rsPen("PE_DB_STATUS")
        End If
    End If
    rsPen.Close
    getWFCPenStatus = retVal
End Function

Public Sub WFC_UptUSBenByEmp(xEmpNo, xBenDate, Optional xEmpHrsWeek = 0, Optional xIsGetHrsWeek = "N", Optional xIsRecalculte = "N", Optional xIsCCDenGTLLOnly = "N", Optional xOldBenGrp, Optional xTranInDate, Optional xRemoveEndDate = "N", Optional xNGSStartDate) 'Ticket #23247 Franks 04/22/2013
Dim rsEmp As New ADODB.Recordset
Dim rsNGS As New ADODB.Recordset
Dim rsBN As New ADODB.Recordset
Dim rsBenT As New ADODB.Recordset
Dim rsBen2 As New ADODB.Recordset
Dim rsEmpPos As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim rsBGMST As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim I, J
Dim xNGSCode, xBenGrpCode, xBenCode, xHrlSal 'xBenDate
Dim xCovAmt, xDIV
Dim xTemp
Dim xNGSStart

On Error GoTo AUDIT_ERR

    If Not glbWFC_US_Ben_Trans Then Exit Sub
    If xIsGetHrsWeek = "N" Then
        If Not IsNumeric(xEmpHrsWeek) Then Exit Sub
        If xEmpHrsWeek < 20 Then Exit Sub
    End If
    
    '''Ticket #23247 Franks 04/09/2014 - begin
    '''If NGS Start Date exist then use it + waiting period as benefit start date
    ''xNGSStart = ""
    ''SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
    ''rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    ''If Not rsEmpOther.EOF Then
    ''    If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
    ''        xNGSStart = rsEmpOther("ER_OTHERDATE1")
    ''    End If
    ''End If
    ''rsEmpOther.Close
    ''If IsDate(xNGSStart) Then
    ''    xBenDate = xNGSStart
    ''End If
    '''Ticket #23247 Franks 04/09/2014 - end
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND LEN(ED_VADIM1 )>0 " 'ED_VADIM1 - NGS Sub Group
    SQLQ = SQLQ & "AND NOT (ED_EMP = 'CB') "
    SQLQ = SQLQ & "AND ED_WORKCOUNTRY = 'U.S.A.' " 'USA employee only
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        'Ticket #25352 Franks 04/16/2014 -
        '"   If Employment Status = COOP or STUD, don't update the Benefit Master Code
        If rsEmp("ED_EMP") = "COOP" Or rsEmp("ED_EMP") = "STUD" Then
            Exit Sub
        End If
        If xIsGetHrsWeek = "Y" Then
            xEmpHrsWeek = 0
            SQLQ = "SELECT JH_EMPNBR,JH_WHRS FROM HR_JOB_HISTORY WHERE JH_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND NOT (JH_CURRENT = 0) "
            SQLQ = SQLQ & "ORDER BY JH_SDATE DESC "
            If rsEmpPos.State <> 0 Then rsEmpPos.Close
            rsEmpPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEmpPos.EOF Then
                If Not IsNull(rsEmpPos("JH_WHRS")) Then
                    xEmpHrsWeek = rsEmpPos("JH_WHRS")
                End If
            End If
            rsEmpPos.Close
            If xEmpHrsWeek < 20 Then Exit Sub
        End If
        If Not IsNull(rsEmp("ED_DIV")) Then xDIV = rsEmp("ED_DIV") Else xDIV = ""
        'If xDiv = "1060" Then GoTo Next_1060
        
        'Hourly or Salaried '
        If IsNull(rsEmp("ED_ORG")) Then GoTo NextRec
        If rsEmp("ED_ORG") = "EXEC" Or rsEmp("ED_ORG") = "NONE" Then
            xHrlSal = "S"
        Else
            xHrlSal = "H"
        End If
        
        'If xIsCCDenGTLLOnly = "Y" Then 'Ticket #24146 Franks 07/26/2013
        '    GoTo Next_1060
        'End If
        
        'find the match NGS Sub Group
        xBenGrpCode = ""
        If xIsRecalculte = "Y" Then 'Find the Benefit Group Code from HREMP
            If Not IsNull(rsEmp("ED_BENEFIT_GROUP")) Then
                xBenGrpCode = rsEmp("ED_BENEFIT_GROUP")
            End If
        Else
            'check with Status
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & rsEmp("ED_DIV") & "' "
            If Not IsNull(rsEmp("ED_ORG")) Then SQLQ = SQLQ & "AND NG_ORG = '" & rsEmp("ED_ORG") & "' "
            SQLQ = SQLQ & "AND NG_PLAN_CODE = '" & rsEmp("ED_EMP") & "' "
            If rsNGS.State <> 0 Then rsNGS.Close
            rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsNGS.EOF Then
                'check "-" status, such as "-ACT2", convert "-ACT2" to "ACT2" with column NEG_STATUS
                SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & rsEmp("ED_DIV") & "' "
                If Not IsNull(rsEmp("ED_ORG")) Then SQLQ = SQLQ & "AND NG_ORG = '" & rsEmp("ED_ORG") & "' "
                SQLQ = SQLQ & "AND LEFT(NG_PLAN_CODE,1) = '-' " 'for "-" code only
                SQLQ = SQLQ & "AND NOT ((CASE LEFT(NG_PLAN_CODE,1) WHEN '-' THEN REPLACE(NG_PLAN_CODE,'-', '') ELSE '' END) = '" & rsEmp("ED_EMP") & "') "
                If rsNGS.State <> 0 Then rsNGS.Close
                rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If rsNGS.EOF Then
                    'No Status
                    SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & rsEmp("ED_DIV") & "' "
                    If Not IsNull(rsEmp("ED_ORG")) Then SQLQ = SQLQ & "AND NG_ORG = '" & rsEmp("ED_ORG") & "' "
                    SQLQ = SQLQ & "AND (NG_PLAN_CODE IS NULL OR NG_PLAN_CODE = '') "
                    If rsNGS.State <> 0 Then rsNGS.Close
                    rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
                End If
            End If
            If Not rsNGS.EOF Then
                If Not IsNull(rsNGS("NG_BENEFIT_GROUP")) Then
                    xBenGrpCode = rsNGS("NG_BENEFIT_GROUP")
                End If
            End If
        End If
        
        
        'Ticket #23247 Franks 09/16/2013 - begin
        If Not IsMissing(xOldBenGrp) And xIsCCDenGTLLOnly = "N" Then
            If Len(xOldBenGrp) > 0 And Len(xBenGrpCode) = 0 Then 'remove BenGrp
                '"   If the Transfer In Division or Union is not a NGS location (ie: not found in the matrix
                '"   Delete the Benefit Group Code
                SQLQ = "UPDATE HREMP SET ED_BENEFIT_GROUP = NULL WHERE ED_EMPNBR = " & xEmpNo & " "
                gdbAdoIhr001.Execute SQLQ
                '"   Add a Benefit End Date to equal the Transfer In Date minus 1.
                xTemp = DateAdd("D", -1, CVDate(xTranInDate))
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND BF_GROUP = '" & xOldBenGrp & "' "
                SQLQ = SQLQ & "AND BF_PCC = 1 "
                'If Not rsBenT.State <> 0 Then rsBenT.Close
                rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                Do While Not rsBenT.EOF
                    rsBenT("BF_CEASEDATE") = CVDate(xTemp)
                    rsBenT.Update
                    'update audit
                    Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
                    rsBenT.MoveNext
                Loop
                rsBenT.Close
                
                'Ticket #24582 Franks 11/08/2013
                'GTLL should have an END DATE because the IE benefit is under $50,000.
                '"   Remove the Benefit Group Code from the Company Paid Benefits
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND BF_BCODE = 'GTLL' "
                SQLQ = SQLQ & "AND NOT (BF_CEASEDATE IS NULL) "
                rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsBenT.EOF Then
                    'check if IE xCovAmt > 50000, if not then enter end date
                    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                    SQLQ = SQLQ & "AND BF_BCODE = 'IE' "
                    If rsBen2.State <> 0 Then rsBen2.Close
                    rsBen2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not rsBen2.EOF Then
                        If Not IsNull(rsBen2("BF_AMT")) Then
                            If rsBen2("BF_AMT") < 50000 Then
                                rsBenT("BF_CEASEDATE") = CVDate(xTemp)
                                rsBenT.Update
                                'update audit
                                Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
                            End If
                        End If
                    End If
                    rsBen2.Close
                End If
                rsBenT.Close
                
                SQLQ = "UPDATE HRBENFT SET BF_GROUP = NULL WHERE BF_EMPNBR = " & xEmpNo & " "
                gdbAdoIhr001.Execute SQLQ
                Exit Sub
            End If
        End If
        'Ticket #23247 Franks 09/16/2013 - end
        
        If Len(xBenGrpCode) = 0 Then GoTo Next_1060 ' NextRec
        
        'If Not rsNGS.EOF Then
            'xBenGrpCode = ""
            'If Not IsNull(rsNGS("NG_BENEFIT_GROUP")) Then
            '    xBenGrpCode = rsNGS("NG_BENEFIT_GROUP")
            'End If

            'benefit screen - begin
            'If Len(xBenGrpCode) = 0 Then GoTo Next_1060 ' NextRec
            SQLQ = "SELECT * FROM HR_BENEFITS_GROUP WHERE BM_BENEFIT_GROUP = '" & xBenGrpCode & "' "
            If rsBGMST.State <> 0 Then rsBGMST.Close
            rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
            Do While Not rsBGMST.EOF
                xBenCode = rsBGMST("BM_BCODE")
                If xEmpHrsWeek >= 0 And xEmpHrsWeek < 30 Then
                    '"   If Hours per Week is less than 30, only add the IE benefit code from the Benefit Group Master
                    If Not xBenCode = "IE" Then
                        GoTo Next_GRP
                    End If
                End If
                
                If xIsCCDenGTLLOnly = "Y" And Not xBenCode = "IE" Then
                    'do nothing for non IE benefits if xIsCCDenGTLLOnly = "Y"
                Else
                    If Not IsMissing(xNGSStartDate) Then 'Ticket #25178 Franks 03/12/2014
                        xCovAmt = WFCUptUSEmpBeneFromBenGRP(xEmpNo, xBenDate, xBenGrpCode, xBenCode, xHrlSal, rsBGMST, xRemoveEndDate, xNGSStartDate)
                    Else
                        xCovAmt = WFCUptUSEmpBeneFromBenGRP(xEmpNo, xBenDate, xBenGrpCode, xBenCode, xHrlSal, rsBGMST, xRemoveEndDate)
                    End If
                End If
                
                '"   If the employee's IE Coverage Amount is greater than $50,000(check Benefit Group to get Amount?), create a Benefit Code GTLL using the logic below
                '"   Create the GTLL with a Coverage of (IE Coverage minus 50000), 100% employee paid, Unit Cost of $2.16 per 1000, Total Annual, Monthly and Pay Period Amount needs to be calculated. The Taxable Benefit flag is set to "Y".
                If xBenCode = "IE" Then
                    If xCovAmt > 50000 Then
                        'create another function to add Benefit which is not based on Ben Group Master
                        'but know CovAmt
                        Call WFCUptUSEmpBeneNotFromBenGRP(xEmpNo, xBenDate, "GTLL", xCovAmt - 50000, 1, 0, 2.16, 1000, xHrlSal, "IE", rsEmp("ED_DOH"))
                    End If
                End If

Next_GRP:
                rsBGMST.MoveNext
            Loop
            rsBGMST.Close
            'benefit screen - end
Next_1060:
            ''If xIsCCDenGTLLOnly = "Y" Then 'Ticket #24146 Franks 07/26/2013
            ''Always check it
                If WFCIsNeedCC(xEmpNo) And xDIV = "1060" Then
                    'Benefit "CC" and "GTLD"
                    'o   If the employee is in Division 1060 and they have a HRDEPEND record with a
                    'Relationship equal to "Son" or "Daughter", create a HRBENFT record with:
                    '"   If the employee's CC Coverage Amount is greater than $2,000, create a Benefit Code GTLD using the logic below:
                    '"   Create the GTLD with a Coverage of (CC Coverage minus 2000), 100% employee paid, Unit Cost of $2.16 per 1000, Total Annual, Monthly and Pay Period Amount needs to be calculated. Refer to the GTLL Pay Period calculation for details. The Taxable Benefit flag is set to "Y".
                    xBenCode = "CC"
                    xCovAmt = 5000
                    '"   Benefit Code = "CC", Coverage = $5,000, 100% company paid, Actual Cost, Total Annual is $8.16, Monthly is .68 and Pay Period Amount is 0.157.
                    Call WFCUptUSEmpBeneNotFromBenGRP(xEmpNo, xBenDate, xBenCode, xCovAmt, 0, 1, 8.16, 0.157, xHrlSal, "IE", rsEmp("ED_DOH"))
                    '"   Create the GTLD with Coverage of $3,000. 100% employee paid, Actual Cost, Total Annual is $8.16, Monthly is .68 and Pay Period Amount is 0.157. The Taxable Benefit flag is set to "Y".
                    'Call WFCUptUSEmpBeneNotFromBenGRP(xEmpNo, xBenDate, "GTLD", xCovAmt - 2000, 1, 0, 8.16, 0.157, xHrlSal, "", rsEmp("ED_DOH"))
                    Call WFCUptUSEmpBeneNotFromBenGRP(xEmpNo, xBenDate, "GTLD", xCovAmt - 2000, 1, 0, 4.896, 0.094, xHrlSal, "CC", rsEmp("ED_DOH"))
                    
                End If
            ''End If
        'End If
        '======================
        
        'Ticket #23247 Franks 09/16/2013 - begin
        '"   If the new Benefit Group does not contain a Benefit that was in the old Benefit Group
        ', the Transfer In Date minus 1 should be the Benefit End Date for that benefit. For example, salaried employee transferring to hourly
        If Not IsMissing(xOldBenGrp) Then
            If Len(xOldBenGrp) > 0 Then  'change BenGrp
                xTemp = DateAdd("D", -1, CVDate(xTranInDate))
                If Not (xOldBenGrp = xBenGrpCode) Then
                    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                    SQLQ = SQLQ & "AND BF_GROUP = '" & xOldBenGrp & "' "
                    ''SQLQ = SQLQ & "AND BF_PCC = 1 "
                    'If Not rsBenT.State <> 0 Then rsBenT.Close
                    rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    Do While Not rsBenT.EOF
                        rsBenT("BF_GROUP") = Null
                        rsBenT("BF_CEASEDATE") = CVDate(xTemp)
                        rsBenT.Update
                        'update audit
                        Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
                        rsBenT.MoveNext
                    Loop
                    rsBenT.Close
                End If
                
                'Ticket #24582 Franks 11/08/2013
                'GTLL should have an END DATE because the IE benefit is under $50,000.
                '"   Remove the Benefit Group Code from the Company Paid Benefits
                SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND BF_BCODE = 'GTLL' "
                SQLQ = SQLQ & "AND  (BF_CEASEDATE IS NULL) "
                rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                If Not rsBenT.EOF Then
                    'check if IE xCovAmt > 50000, if not then enter end date
                    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
                    SQLQ = SQLQ & "AND BF_BCODE = 'IE' "
                    If rsBen2.State <> 0 Then rsBen2.Close
                    rsBen2.Open SQLQ, gdbAdoIhr001, adOpenStatic
                    If Not rsBen2.EOF Then
                        If Not IsNull(rsBen2("BF_AMT")) Then
                            If rsBen2("BF_AMT") < 50000 Then
                                rsBenT("BF_CEASEDATE") = CVDate(xTemp)
                                rsBenT.Update
                                'update audit
                                Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
                            End If
                        End If
                    End If
                    rsBen2.Close
                End If
                rsBenT.Close
                
            End If
        End If
        'Ticket #23247 Franks 09/16/2013 - end
    End If
    rsEmp.Close
    
NextRec:

Exit Sub
AUDIT_ERR:
Debug.Print Err.Description
Resume Next

End Sub


Public Function WFCUptUSEmpBeneFromBenGRP(xEmpNo, xBenDate, xBenGrpCode, xBenCode, xHrlSal, rsBGMST As ADODB.Recordset, Optional xRemoveEndDate = "N", Optional xNGSStartDate) 'Ticket #23247 Franks 04/22/2013
Dim rsLocEmp As New ADODB.Recordset
Dim rsBN As New ADODB.Recordset
Dim SQLQ  As String
Dim I, J, K, M, totNum
Dim xSalFactor, xRndFactor, xMaxCover, xMinCover, xCovAmount, AnCoverAmt
Dim xSalary
Dim CostINFO As BenefitCost
Dim xDecimal
Dim xWaitPeriod, xDWM, xNEXTNEAREST, xLoBenDate
Dim xIsSalDep As String
Dim NomalCCost
Dim AMT, tmpTotalCost As Double, xPER, xUNITCOST
Dim xFinBenDate
Dim OBCode, OPPAMT, OMTHCOMP, OMTHEMP, OBAMT, OPPE, OPCC, OMAXDOL, OEDate, OCOVER, OTCOST
Dim OBenGrp, OCeasedate 'Ticket #23247 Franks 09/16/2013
Dim OWaitPeriod  'Ticket #24073 07/16/2013
Dim xNewRec As Boolean
Dim retVal 'return Coverage Amount
Dim xPPAMT
On Error GoTo AUDIT_ERR

    retVal = 0
    WFCUptUSEmpBeneFromBenGRP = 0
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsLocEmp.EOF Then Exit Function
    
    'get Effective Date from Ben Grp Master
    xFinBenDate = xBenDate
    If Not IsMissing(xNGSStartDate) Then 'Ticket #25178 Franks 03/12/2014
        xLoBenDate = getBenEDateFromWaitingPeriod(xEmpNo, xNGSStartDate, rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"))
    Else
        xLoBenDate = getBenEDateFromWaitingPeriod(xEmpNo, rsLocEmp("ED_DOH"), rsBGMST("BM_WAITPERIOD"), rsBGMST("BM_DWM"))
    End If
    If IsDate(xFinBenDate) And IsDate(xLoBenDate) Then 'new hire may have effective date later than 01/01/2013
        If CVDate(xLoBenDate) > CVDate(xFinBenDate) Then
            xFinBenDate = xLoBenDate
        End If
    End If
    
    'If xBenCode = "SD" Then
    '    Debug.Print ""
    'End If
    'if Salary Dependent
    xIsSalDep = "N"
    If Not IsNull(rsBGMST("BM_SALARYDEPENDANT")) Then
        xIsSalDep = rsBGMST("BM_SALARYDEPENDANT") 'Y N
    End If
    
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND BF_BCODE = '" & xBenCode & "' "
    If rsBN.State <> 0 Then rsBN.Close
    rsBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsBN.EOF Then xNewRec = True Else xNewRec = False
    'Call setBenOldVal4Audit(rsBN, xNewRec, OBCode, OPPAMT, OBAMT, OEDate, OCOVER, OTCOST)
    Call setBenOldVal4Audit(rsBN, xNewRec, OBCode, OPPAMT, OBAMT, OEDate, OCOVER, OTCOST, OBenGrp, OCeasedate, OWaitPeriod)
    
    'Ticket #23247 Franks 04/25/2013 - begin
    If IsNull(rsBGMST("BM_FACTOR")) Then xSalFactor = 1 Else xSalFactor = rsBGMST("BM_FACTOR") ' Val(medSalFactor)
    If IsNull(rsBGMST("BM_ROUND")) Then xRndFactor = 0 Else xRndFactor = rsBGMST("BM_ROUND") 'xRndFactor = Val(txtRoundFactor)
    If IsNull(rsBGMST("BM_MAXIMUM")) Then xMaxCover = 0 Else xMaxCover = rsBGMST("BM_MAXIMUM") 'xMaxCover = Val(medMaxCover)
    If IsNull(rsBGMST("BM_MINIMUM")) Then xMinCover = 0 Else xMinCover = rsBGMST("BM_MINIMUM") 'xMinCover = Val(medMinCover)
    If IsNull(rsBGMST("BM_PER")) Then xPER = 0 Else xPER = rsBGMST("BM_PER")
    If IsNull(rsBGMST("BM_UNITCOST")) Then xUNITCOST = 0 Else xUNITCOST = rsBGMST("BM_UNITCOST")
    If IsNull(rsBGMST("BM_WaitPeriod")) Then xWaitPeriod = "" Else xWaitPeriod = rsBGMST("BM_WaitPeriod") 'Ticket #24073 07/16/2013
          
    xPPAMT = 0
    xCovAmount = 0
    tmpTotalCost = 0
    'Coverage Amount - begin
    If xIsSalDep = "Y" Then
        xSalary = CrtSalary(xEmpNo)
        CostINFO = CrtBeneCost(xEmpNo, xSalary, xBenGrpCode, xBenCode)
        xSalary = CostINFO.Salary
        xCovAmount = xSalary * xSalFactor
        'rounding moved before setting to min or max by Bryan 18/Oct/05 Ticket#9487
        If xRndFactor = 0 Then xRndFactor = 0.01
        
        If IsNull(rsBGMST("BM_NEXTNEAREST")) Then xNEXTNEAREST = "N" Else xNEXTNEAREST = rsBGMST("BM_NEXTNEAREST")
        'If Rounding factor is 1 and Next, and Coverage Amount is a whole # then do not add 0.5
        If xRndFactor = 1 And xNEXTNEAREST = "N" Then ' optRound(0) = False Then
            xDecimal = xCovAmount - Int(xCovAmount)
            If xDecimal = 0 Then
                xCovAmount = Round(xCovAmount / xRndFactor) * xRndFactor
            Else
                xCovAmount = Round(xCovAmount / xRndFactor + IIf(xNEXTNEAREST = "R", 0, 0.5)) * xRndFactor
            End If
        Else
            'Ticket #18465 - If NEXT and evenly divisible (whole #) then do not round to NEXT
            If xNEXTNEAREST = "N" Then 'optRound(0) = False Then
                xDecimal = (xCovAmount / xRndFactor) - Int(xCovAmount / xRndFactor)
                If xDecimal = 0 Then
                    xCovAmount = Round(xCovAmount / xRndFactor) * xRndFactor
                Else
                    xCovAmount = Round(xCovAmount / xRndFactor + IIf(xNEXTNEAREST = "R", 0, 0.5)) * xRndFactor
                End If
            Else
                xCovAmount = Round(xCovAmount / xRndFactor + IIf(xNEXTNEAREST = "R", 0, 0.5)) * xRndFactor
            End If
        End If
        
        If xMinCover <> 0 And xCovAmount < xMinCover Then xCovAmount = xMinCover
        If xMaxCover <> 0 And xCovAmount > xMaxCover Then xCovAmount = xMaxCover
        
        'medCovAmount = xCovAmount
        'rsBN("BF_AMT") = xCovAmount
        If CostINFO.Type = "M" Or CostINFO.Type = "W" Then  'Ticket #25235 - For weekly too * 12 even though the Covrg Amt is Weekly
            AnCoverAmt = xCovAmount * 12
        'Ticket #25235 - This is not working so Jerry and I decided to use the above * 12 which gives the right result for the client
        'ElseIf CostINFO.Type = "W" Then     'Ticket #22682 - Release 8.0 - added Weekly option to Benefit Costing
        '    AnCoverAmt = xCovAmount * 52
        Else
            AnCoverAmt = xCovAmount
        End If
    End If
    If xIsSalDep = "N" Then
        If Not rsBN.EOF Then
            rsBN("BF_EDATE") = CVDate(xFinBenDate) 'Ticket #23247 Franks 09/16/2013
            rsBN("BF_GROUP") = xBenGrpCode 'Ticket #23247 Franks 09/16/2013
        End If
        xCovAmount = rsBGMST("BM_AMT")
        'rsBN("BF_AMT") = rsBGMST("BM_AMT")
    End If
    'Coverage Amount - end
    
    'Total - begin ---------------------------------------------
    'Call setTotal()
    tmpTotalCost = 0
    If Not IsNull(rsBGMST("BM_TCOST")) Then
        tmpTotalCost = rsBGMST("BM_TCOST") 'for Non Salary Dependent Benefit
    End If
    If xIsSalDep = "N" Then  ' comSalDepn = "Yes" Then
        AMT = xCovAmount 'AnCoverAmt
    Else
        AMT = xCovAmount
    End If
    If rsBGMST("BM_PREMIUM") = "P" Then 'optActual(1) = True
        'If Val(txtPer) > 0 Then tmpTotalCost = (Val(AMT) / Val(txtPer)) * Val(medUnitCost) Else tmpTotalCost = 0
        If xPER > 0 Then tmpTotalCost = (AMT / xPER) * xUNITCOST Else tmpTotalCost = 0
    End If
    'Total - end   ---------------------------------------------
    
    '"   The Pay Period Amount will need to be populated based on
    '"   Hourly: Total Company Cost / 52 (rounded to 2 decimals)
    '"   Salaried: Total Company Cost / 24 (rounded to 2 decimals)
    'rsBN("BF_PPAMT") = rsBGMST("BM_PPAMT")
    If IsNumeric(tmpTotalCost) Then 'Not IsNull(rsBGMST("BM_TCOST")) Then
        If xHrlSal = "H" Then
            xPPAMT = Round((tmpTotalCost / 52), 2)
        Else 'Salaried
            xPPAMT = Round((tmpTotalCost / 24), 2)
        End If 'BF_TCOST
    End If

    If Not xNewRec Then 'for update
        If IsNumeric(OPPAMT) And IsNumeric(OBAMT) And IsNumeric(OTCOST) Then
            ''If Round(OPPAMT, 2) = Round(xPPAMT, 2) And Round(OBAMT, 2) = Round(xCovAmount, 2) And Round(OTCOST, 2) = Round(tmpTotalCost, 2) Then
            ''Ticket #24073 07/16/2013 - added OWaitPeriod
            ''If Round(OPPAMT, 2) = Round(xPPAMT, 2) And Round(OBAMT, 2) = Round(xCovAmount, 2) And Round(OTCOST, 2) = Round(tmpTotalCost, 2) And OWaitPeriod = xWaitPeriod Then
            'Ticket #23247 Franks 09/16/2013
            If Round(OPPAMT, 2) = Round(xPPAMT, 2) And Round(OBAMT, 2) = Round(xCovAmount, 2) And Round(OTCOST, 2) = Round(tmpTotalCost, 2) And OWaitPeriod = xWaitPeriod And OBenGrp = xBenGrpCode Then ' And OEDate = xFinBenDate Then
                If IsNull(rsBN("BF_AMT")) Then retVal = 0 Else retVal = rsBN("BF_AMT") 'Ticket #23247 Franks 04/09/2014 for GTLL calculation
                WFCUptUSEmpBeneFromBenGRP = retVal
                Exit Function 'no change then exit - for Recalculate on Benefit/Benefit Group
            End If
        End If
    End If
    'Ticket #23247 Franks 04/25/2013 - end
    
    If rsBN.EOF Then
        rsBN.AddNew
        rsBN("BF_EMPNBR") = xEmpNo
        'rsBN("BF_EDATE") = CVDate(xFinBenDate)
    Else
        If xRemoveEndDate = "Y" Then 'Ticket #24451 Franks 10/17/2013
            rsBN("BF_CEASEDATE") = Null
        End If
    End If
    rsBN("BF_GROUP") = xBenGrpCode
    rsBN("BF_BCODE") = xBenCode
    rsBN("BF_EDATE") = CVDate(xFinBenDate)
    'Coverage Amount - begin '
    rsBN("BF_AMT") = xCovAmount
    retVal = rsBN("BF_AMT")
    'Coverage Amount - end
    
    'Total - begin
    If Not IsNull(rsBGMST("BM_PCE")) Then
        'medEECost = Val(tmpTotalCost) * Val(medPPE)
        rsBN("BF_ECOST") = tmpTotalCost * rsBGMST("BM_PCE")
    Else
        rsBN("BF_ECOST") = 0
    End If
    If Not IsNull(rsBGMST("BM_PCC")) Then
        'medCompCost = Val(tmpTotalCost) * Val(medPPComp)
        rsBN("BF_CCOST") = tmpTotalCost * rsBGMST("BM_PCC")
    Else
        rsBN("BF_CCOST") = 0
    End If
    rsBN("BF_MTHCCOST") = rsBN("BF_CCOST") / 12
    rsBN("BF_MTHECOST") = rsBN("BF_ECOST") / 12
    rsBN("BF_TCOST") = tmpTotalCost
    'Total - end
    
    '"   The Pay Period Amount will need to be populated based on
    '"   Hourly: Total Company Cost / 52 (rounded to 2 decimals)
    '"   Salaried: Total Company Cost / 24 (rounded to 2 decimals)
    rsBN("BF_PPAMT") = xPPAMT  'rsBGMST("BM_PPAMT")
    rsBN("BF_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
    rsBN("BF_DWM") = rsBGMST("BM_DWM")
    rsBN("BF_COVER") = rsBGMST("BM_COVER")
    'rsBN("BF_AMT") = rsBGMST("BM_AMT")
    rsBN("BF_UNITCOST") = rsBGMST("BM_UNITCOST")
    rsBN("BF_PCE") = rsBGMST("BM_PCE")
    rsBN("BF_PCC") = rsBGMST("BM_PCC")
    'rsBN("BF_ECOST") = rsBGMST("BM_ECOST")
    'rsBN("BF_CCOST") = rsBGMST("BM_CCOST")
    'rsBN("BF_TCOST") = rsBGMST("BM_TCOST")
    rsBN("BF_MAXDOL") = rsBGMST("BM_MAXDOL")
    rsBN("BF_PREMIUM") = rsBGMST("BM_PREMIUM")
    rsBN("BF_PER") = rsBGMST("BM_PER")
    'rsBN("BF_MTHCCOST") = rsBGMST("BM_MTHCCOST")
    'rsBN("BF_MTHECOST") = rsBGMST("BM_MTHECOST")
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
    rsBN("BF_RATELEVEL") = rsBGMST("BM_RATELEVEL")
    rsBN("BF_LUSER") = glbUserID
    rsBN("BF_LDATE") = Date
    rsBN("BF_LTIME") = Time$
    rsBN.Update
    
    'update HRAUDIT table
    Call WFC_AUDITBENF(xEmpNo, xNewRec, rsBN, OBCode, OPPAMT, OBAMT, OEDate, OCOVER, OTCOST)
    
    rsBN.Close
                
    WFCUptUSEmpBeneFromBenGRP = retVal
Exit Function
AUDIT_ERR:
Debug.Print Err.Description
Resume Next

End Function
Public Function WFCIsNeedCC(xEmpNo) 'Ticket #23247 Franks 04/22/2013
Dim rsLocEmp As New ADODB.Recordset
Dim rsDep As New ADODB.Recordset
Dim SQLQ  As String
Dim retVal As Boolean
'o   If the employee is in Division 1060 and the Union Code is U959 and they have a HRDEPEND record with a
'Relationship equal to "Son" or "Daughter", create a HRBENFT record with:
    retVal = False
    SQLQ = "SELECT ED_EMPNBR, ED_DIV FROM HREMP WHERE ED_DIV = '1060' AND ED_ORG = 'U959' AND ED_EMPNBR = " & xEmpNo & " "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLocEmp.EOF Then
        ''    '''If Not IsNull(rsLocEmP("ED_DIV")) Then
        ''    ''    'If rsLocEmP("ED_DIV") = "1060" Then
        ''    ''        SQLQ = "SELECT DP_EMPNBR,DP_RELATE FROM HRDEPEND WHERE DP_EMPNBR = " & xEmpNo & " "
        ''    ''        'SQLQ = SQLQ & "AND (DP_RELATE = 'Son' OR DP_RELATE = 'Daughter') "
        ''    ''        If rsDep.State <> 0 Then rsDep.Close
        ''    ''        rsDep.Open SQLQ, gdbAdoIhr001, adOpenStatic
        ''    ''        If Not rsDep.EOF Then
        ''    ''            retVal = True
        ''    ''        End If
        ''    ''        rsDep.Close
        ''    ''    'End If
        ''    '''End If
        ''
        ''    'Ticket #23247 Franks 07/22/2013
        ''    'Jerry needs these two benefits(CC and GTLD) to be setup on new hire, so we don't use this any more
        ''    '    Relationship equal to "Son" or "Daughter", create a HRBENFT record with:
        ''    retVal = True
        
        'Ticket #24146 Franks 07/26/2013
        'dont check dependent table
        retVal = True
    End If
    WFCIsNeedCC = retVal
End Function

Public Sub WFCUptUSEmpBeneNotFromBenGRP(xEmpNo, xBenDate, xBenCode, xCovAmt, xPercE, xPercC, xUNITCOST, xPER, xHrlSal, xBenRelated, xDOH, Optional xIsRecalculte = "N") 'Ticket #23247 Franks 04/22/2013
Dim rsLocEmp As New ADODB.Recordset
Dim rsBN As New ADODB.Recordset
Dim xLoBenDate
Dim xFinBenDate
Dim SQLQ  As String
Dim tmpTotalCost
Dim OBCode, OPPAMT, OMTHCOMP, OMTHEMP, OBAMT, OPPE, OPCC, OMAXDOL, OEDate, OCOVER, OTCOST
Dim OBenGrp, OCeasedate 'Ticket #23247 Franks 09/16/2013
Dim xNewRec As Boolean
Dim xPPAMT
'Note for Frank: there is a same function in IHREI, any change must be done in IHREI as well

On Error GoTo AUDIT_ERR

    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsLocEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsLocEmp.EOF Then Exit Sub
    
    
    'get BenDate from Related Benefit Code
    xFinBenDate = xBenDate
    xLoBenDate = ""
    If Len(xBenRelated) > 0 Then
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_BCODE = '" & xBenRelated & "' "
        SQLQ = SQLQ & "ORDER BY BF_EDATE DESC "
        If rsBN.State <> 0 Then rsBN.Close
        rsBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsBN.EOF Then
            xLoBenDate = rsBN("BF_EDATE")
        End If
        rsBN.Close
    Else
        'compare DOH
        If IsDate(xDOH) Then
            xLoBenDate = xDOH
        End If
    End If
    If IsDate(xFinBenDate) And IsDate(xLoBenDate) Then 'new hire may have effective date later than 01/01/2013
        If CVDate(xLoBenDate) > CVDate(xFinBenDate) Then
            xFinBenDate = xLoBenDate
        End If
    End If
    
    'Add or update Benefit - begin
    SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND BF_BCODE = '" & xBenCode & "' "
    If rsBN.State <> 0 Then rsBN.Close
    rsBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If rsBN.EOF Then xNewRec = True Else xNewRec = False
    Call setBenOldVal4Audit(rsBN, xNewRec, OBCode, OPPAMT, OBAMT, OEDate, OCOVER, OTCOST, OBenGrp, OCeasedate)
    
    
    'Ticket #23247 Franks 04/25/2013 - begin
    'Total & PPAMT - begin
    tmpTotalCost = 0
    xPPAMT = 0
    If xBenCode = "CC" Or xBenCode = "GTLD" Then
        tmpTotalCost = xUNITCOST
        xPPAMT = xPER
    Else
        If xPER > 0 Then tmpTotalCost = (xCovAmt / xPER) * xUNITCOST Else tmpTotalCost = 0
        '"   The Pay Period Amount will need to be populated based on
        '"   Hourly: Total Company Cost / 52 (rounded to 2 decimals)
        '"   Salaried: Total Company Cost / 24 (rounded to 2 decimals)
        'rsBN("BF_PPAMT") = rsBGMST("BM_PPAMT")
        If IsNumeric(tmpTotalCost) Then 'Not IsNull(rsBGMST("BM_TCOST")) Then
            If xHrlSal = "H" Then
                xPPAMT = Round((tmpTotalCost / 52), 2)
            Else 'Salaried
                xPPAMT = Round((tmpTotalCost / 24), 2)
            End If 'BF_TCOST
        End If
    End If
    'Total & PPAMT - end
    
    If Not xNewRec Then 'for update
        'If IsNumeric(OBAMT) And IsNumeric(OTCOST) Then
        If IsNumeric(OPPAMT) And IsNumeric(OBAMT) And IsNumeric(OTCOST) Then
            If Round(OPPAMT, 2) = Round(xPPAMT, 2) And Round(OBAMT, 2) = Round(xCovAmt, 2) And Round(OTCOST, 2) = Round(tmpTotalCost, 2) Then
            ''PPAMT was changed on the Regular Benefit Group calculation every time, so don't use it
            'If Round(OBAMT, 2) = Round(xCovAmount, 2) And Round(OTCOST, 2) = Round(tmpTotalCost, 2) Then
                Exit Sub 'no change then exit - for Recalculate on Benefit/Benefit Group
            End If
        End If
    End If
    'Ticket #23247 Franks 04/25/2013 - end
    
    If rsBN.EOF Then
        rsBN.AddNew
        rsBN("BF_EMPNBR") = xEmpNo
    End If
    'rsBN("BF_GROUP") = xBenGrpCode
    rsBN("BF_BCODE") = xBenCode
    If xIsRecalculte = "Y" And IsDate(rsBN("BF_EDATE")) Then
        'don't change the BF_EDATE on Recalculte
    Else
        rsBN("BF_EDATE") = CVDate(xFinBenDate)
    End If
    
    'If IsNull(rsBGMST("BM_FACTOR")) Then xSalFactor = 1 Else xSalFactor = rsBGMST("BM_FACTOR") ' Val(medSalFactor)
    'If IsNull(rsBGMST("BM_ROUND")) Then xRndFactor = 0 Else xRndFactor = rsBGMST("BM_ROUND") 'xRndFactor = Val(txtRoundFactor)
    'If IsNull(rsBGMST("BM_MAXIMUM")) Then xMaxCover = 0 Else xMaxCover = rsBGMST("BM_MAXIMUM") 'xMaxCover = Val(medMaxCover)
    'If IsNull(rsBGMST("BM_MINIMUM")) Then xMinCover = 0 Else xMinCover = rsBGMST("BM_MINIMUM") 'xMinCover = Val(medMinCover)
    'If IsNull(rsBGMST("BM_PER")) Then xPER = 0 Else xPER = rsBGMST("BM_PER")
    'If IsNull(rsBGMST("BM_UNITCOST")) Then xUnitCost = 0 Else xUnitCost = rsBGMST("BM_UNITCOST")
    rsBN("BF_AMT") = xCovAmt
    rsBN("BF_PCE") = xPercE
    rsBN("BF_PCC") = xPercC
    
    If xBenCode = "CC" Or xBenCode = "GTLD" Then
        rsBN("BF_PREMIUM") = "A" 'Actual
    Else
        rsBN("BF_UNITCOST") = xUNITCOST
        rsBN("BF_PER") = xPER
        rsBN("BF_PREMIUM") = "P" ' "A" 'Actual
    End If
    rsBN("BF_PPAMT") = xPPAMT
    
    'If Not IsNull(rsBGMST("BM_PCE")) Then
    If IsNumeric(xPercE) Then
        'medEECost = Val(tmpTotalCost) * Val(medPPE) '
        rsBN("BF_ECOST") = tmpTotalCost * xPercE 'rsBGMST("BM_PCE")
    Else
        rsBN("BF_ECOST") = 0
    End If
    'If Not IsNull(rsBGMST("BM_PCC")) Then
    If IsNumeric(xPercC) Then
        'medCompCost = Val(tmpTotalCost) * Val(medPPComp)
        rsBN("BF_CCOST") = tmpTotalCost * xPercC 'rsBGMST("BM_PCC")
    Else
        rsBN("BF_CCOST") = 0
    End If
    rsBN("BF_MTHCCOST") = rsBN("BF_CCOST") / 12
    rsBN("BF_MTHECOST") = rsBN("BF_ECOST") / 12
    rsBN("BF_TCOST") = tmpTotalCost
    'Total - end
    
    If xBenCode = "GTLL" Or xBenCode = "GTLD" Then
        rsBN("BF_TAXBEN") = "Y"
    End If
    
    'rsBN("BF_WAITPERIOD") = rsBGMST("BM_WAITPERIOD")
    'rsBN("BF_DWM") = rsBGMST("BM_DWM")
    'rsBN("BF_COVER") = rsBGMST("BM_COVER")
    'rsBN("BF_AMT") = rsBGMST("BM_AMT")
    'rsBN("BF_MAXDOL") = rsBGMST("BM_MAXDOL")
    'rsBN("BF_PREMIUM") = rsBGMST("BM_PREMIUM")
    'rsBN("BF_PER") = rsBGMST("BM_PER")
    'rsBN("BF_MTHCCOST") = rsBGMST("BM_MTHCCOST")
    'rsBN("BF_MTHECOST") = rsBGMST("BM_MTHECOST")
    'rsBN("BF_SALARYDEPENDANT") = rsBGMST("BM_SALARYDEPENDANT")
    'rsBN("BF_MINIMUM") = rsBGMST("BM_MINIMUM")
    'rsBN("BF_FACTOR") = rsBGMST("BM_FACTOR")
    'rsBN("BF_ROUND") = rsBGMST("BM_ROUND")
    'rsBN("BF_MAXIMUM") = rsBGMST("BM_MAXIMUM")
    'rsBN("BF_NEXTNEAREST") = rsBGMST("BM_NEXTNEAREST")
    'rsBN("BF_TAXAMOUNT") = rsBGMST("BM_TAXAMOUNT")
    'rsBN("BF_COMMENTS") = rsBGMST("BM_COMMENTS")
    'rsBN("BF_PTAX") = rsBGMST("BM_PTAX")
    'rsBN("BF_PERORDOLL") = rsBGMST("BM_PERORDOLL")
    'rsBN("BF_POLICY") = rsBGMST("BM_POLICY") 'Ticket #13448 WFC Manulife needs Policy Number
    'rsBN("BF_RATELEVEL") = rsBGMST("BM_RATELEVEL")
    rsBN("BF_LUSER") = glbUserID
    rsBN("BF_LDATE") = Date
    rsBN("BF_LTIME") = Time$
    rsBN.Update
    
    'update HRAUDIT table
    Call WFC_AUDITBENF(xEmpNo, xNewRec, rsBN, OBCode, OPPAMT, OBAMT, OEDate, OCOVER, OTCOST)
    
    rsBN.Close
    'Add or update Benefit - end
Exit Sub
AUDIT_ERR:
Debug.Print Err.Description
Resume Next

End Sub


Public Function getBenEDateFromWaitingPeriod(xEmpNo, xDOH, xWaitPeriod, xDWM)
'Dim rsEmp As New ADODB.Recordset
Dim xPER
Dim retVal
    retVal = xDOH
    If IsNumeric(xWaitPeriod) Then
        'rsEmp.Open "SELECT ED_DOH, ED_USRDAT1 FROM HREMP WHERE ED_EMPNBR=" & xEmpNo, gdbAdoIhr001, adOpenStatic
        'If Not rsEmp.EOF Then
            'If IsDate(rsEmp("ED_DOH")) Then
            If IsDate(xDOH) Then
                If Len(xDWM) > 0 Then
                    xPER = "M"
                    If xDWM <> "" Then
                       xPER = Left(xDWM, 1) 'D, W, M
                       If xPER = "W" Then xPER = "WW"
                    End If
                    retVal = DateAdd(xPER, Val(xWaitPeriod), xDOH) 'rsEmp("ED_DOH"))
                End If
            End If
            'End If
        'End If
    End If
    getBenEDateFromWaitingPeriod = retVal
End Function

Public Sub setBenOldVal4Audit(rslBen As ADODB.Recordset, xlocNewRec As Boolean, OBCode, OPPAMT, OBAMT, OEDate, OCOVER, OTCOST, OBenGrp, OCeasedate, Optional OWaitPeriod = "")
    OBCode = ""
    OPPAMT = ""
    'OMTHCOMP = ""
    'OMTHEMP = ""
    OBAMT = ""
    'OPPE = ""
    'OPCC = ""
    'OMAXDOL = ""
    OEDate = ""
    OCOVER = ""
    OTCOST = ""
    OBenGrp = ""
    OCeasedate = ""
    If Not IsMissing(OWaitPeriod) Then 'Ticket #24073 Franks 07/16/2013
        OWaitPeriod = ""
    End If
    If Not xlocNewRec And Not rslBen.EOF Then
        OBCode = rslBen("BF_BCODE")
        If Not IsNull(rslBen("BF_PPAMT")) Then OPPAMT = rslBen("BF_PPAMT")
        'If Not IsNull(rslBen("BF_MTHCCOST")) Then OMTHCOMP = rslBen("BF_MTHCCOST")
        'If Not IsNull(rslBen("BF_MTHECOST")) Then OMTHEMP = rslBen("BF_MTHECOST")
        If Not IsNull(rslBen("BF_AMT")) Then OBAMT = rslBen("BF_AMT")
        'If Not IsNull(rslBen("BF_PCC")) Then OPCC = rslBen("BF_PCC")
        'If Not IsNull(rslBen("BF_PCE")) Then OPPE = rslBen("BF_PCE")
        'If Not IsNull(rslBen("BF_MAXDOL")) Then OMAXDOL = rslBen("BF_MAXDOL")
        If Not IsNull(rslBen("BF_EDATE")) Then OEDate = rslBen("BF_EDATE")
        If Not IsNull(rslBen("BF_COVER")) Then OCOVER = rslBen("BF_COVER")
        If Not IsNull(rslBen("BF_TCOST")) Then OTCOST = rslBen("BF_TCOST")
        If Not IsMissing(OWaitPeriod) Then 'Ticket #24073 Franks 07/16/2013
            If Not IsNull(rslBen("BF_WaitPeriod")) Then OWaitPeriod = rslBen("BF_WaitPeriod") Else OWaitPeriod = ""
        End If
        If Not IsNull(rslBen("BF_GROUP")) Then OBenGrp = rslBen("BF_GROUP")
        If Not IsNull(rslBen("BF_CEASEDATE")) Then OCeasedate = rslBen("BF_CEASEDATE")
    End If
End Sub

Public Function WFC_AUDITBEN_ByField(xEmpNo, ACTX, xFieldName, rsBenT As ADODB.Recordset)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
Dim rsEmp As New ADODB.Recordset
Dim SQLQ
    
'''On Error GoTo AUDIT_ERR
WFC_AUDITBEN_ByField = False

rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDIV = ""
    Else
        xDIV = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDIV = ""
End If
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_PER, AU_BAMT, AU_UNITCOST,AU_CEASEDATE, "
strFields = strFields & "AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

'SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
'SQLQ = SQLQ & "AND BF_BCODE = '" & xBenCode & "' "
'rsBenT.Open SQLQ, gdbAdoIhr001, adOpenStatic
'If Not rsBenT.EOF Then
'    GoTo End_Line
'End If

xADD = False

If ACTX = "D" Then GoTo MODUPD
If ACTX = "M" Then GoTo MODNOUPD

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDIV

If ACTX = "D" Then
  'If aType = 1 Then
    rsTA("AU_BCODE") = rsBenT("BF_BCODE") ' clpCode(1).Text
    'If txtCovType <> "" Then rsTA("AU_COVER") = txtCovType
    rsTA("AU_COVER") = rsBenT("BF_COVER")
    rsTA("AU_EDATE") = rsBenT("BF_EDATE") 'dlpDate(0).Text
    rsTA("AU_MAXDOL") = rsBenT("BF_MAXDOL") ' Val(medMaxAmnt)
    rsTA("AU_PPAMT") = rsBenT("BF_PPAMT") '  medPayPeriodAmount
    rsTA("AU_MTHCCOST") = rsBenT("BF_MTHCCOST") ' medMCCOST
    rsTA("AU_MTHECOST") = rsBenT("BF_MTHECOST") ' medMECOST
  'End If
End If

SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
rsTA("AU_LDATE") = Date 'today
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update

GoTo end_line

MODNOUPD:

'Ticket #23247 Franks 09/16/2013
If Not xFieldName = "BF_CEASEDATE" Then 'for BF_CEASEDATE update
    GoTo end_line
End If

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDIV


rsTA("AU_BCODE") = rsBenT("BF_BCODE") ' clpCode(1).Text
rsTA("AU_CEASEDATE") = rsBenT("BF_CEASEDATE")
'If txtCovType <> "" Then rsTA("AU_COVER") = txtCovType
'rsTA("AU_EDATE") = rsBenT("BF_EDATE") 'dlpDate(0).Text
'rsTA("AU_MAXDOL") = rsBenT("BF_MAXDOL") ' Val(medMaxAmnt)
'rsTA("AU_PPAMT") = rsBenT("BF_PPAMT") '  medPayPeriodAmount
'rsTA("AU_MTHCCOST") = rsBenT("BF_MTHCCOST") ' medMCCOST
'rsTA("AU_MTHECOST") = rsBenT("BF_MTHECOST") ' medMECOST

SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close

rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
If IsNull(rsBenT("BF_CEASEDATE")) Then
    rsTA("AU_LDATE") = Date 'today
Else
    If CVDate(rsBenT("BF_CEASEDATE")) > CVDate(Date) Then
        rsTA("AU_LDATE") = rsBenT("BF_CEASEDATE")
    Else
        rsTA("AU_LDATE") = Date '
    End If
End If
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update


end_line:

WFC_AUDITBEN_ByField = True
Exit Function
AUDIT_ERR:

End Function

Public Function WFC_AUDITEMPPOS(xEmpNo, xJob, xDATE) 'Ticket #29183 Franks 09/13/2016
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD, xPT, xDIV
Dim HRChangs As New Collection
Dim HRSalary As New Collection
Dim UpdateAudit As Boolean
Dim UptPositionDate As Date
Dim HRChangs1 As New Collection
Dim strFields As String
Dim rsEmp As New ADODB.Recordset
Dim SQLQ

On Error GoTo AUDIT_ERR

AUDITPSTN = False

If rsTB.State <> 0 Then rsTB.Close
rsTB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then xPT = "" Else xPT = rsTB("ED_PT")
    If IsNull(rsTB("ED_DIV")) Then xDIV = "" Else xDIV = rsTB("ED_DIV")
Else
    xPT = ""
    xDIV = ""
End If

strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_PHRS, AU_OLDPHRS, AU_WHRS, AU_OLDWHRS, AU_DHRS, AU_OLDDHRS, "
strFields = strFields & "AU_JOB, AU_SJDATE, AU_JREASON, AU_LEADHAND, AU_LABOURCD, AU_LABOUREDATE, "
strFields = strFields & "AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID, AU_ORG, AU_BILLINGRATE "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic

xADD = False

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP": rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM": rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDIV
rsTA("AU_JOB") = xJob
rsTA("AU_SJDATE") = xDATE
'rsTA("AU_JREASON") = clpCode(1).Text
''If oPHRS <> medHours(2) Then
''    rsTA("AU_PHRS") = medHours(2)
''    rsTA("AU_OLDPHRS") = oPHRS
''End If
''If oWHRS <> medHours(1) Then
''    rsTA("AU_WHRS") = medHours(1)
''    rsTA("AU_OLDWHRS") = oWHRS
''End If
''If ODHRS <> medHours(0) Then
''    If Not IsNumeric(medHours(0)) Then medHours(0) = 0
''    rsTA("AU_DHRS") = medHours(0)
''    rsTA("AU_OLDDHRS") = ODHRS
''End If
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
If CVDate(xDATE) > CVDate(Date) Then
    rsTA("AU_LDATE") = CVDate(xDATE)
Else
    rsTA("AU_LDATE") = Date
End If
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = "A" 'ACTX
SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
rsTA.Update
rsTA.Close


MODNOUPD:
AUDITPSTN = True

Exit Function

AUDIT_ERR:
MsgBox Err.Description
Exit Function
Resume Next

End Function

Public Sub Upd_Related_Salary_public(xEmpNo, rsEPos As ADODB.Recordset, xIsUptAuditBen)
Dim SQLQ As String, Msg As String
Dim dynHRSALHIS As New ADODB.Recordset
Dim JobCode$, PositionStartDat, JobReason$
Dim HoursPerWeek!
Dim lngJobID&
Dim X!, cX$
Dim SH_SALARY@, SH_SALCD$, SH_EDATE, SH_PAYP$, SH_NEXTDAT As Variant
Dim xSH_FISCALYEAR, xSH_SECTION, xSH_MARKETLINE, xSH_BAND, xPayrollID, xSH_CURRENCYINDI 'WFC ONLY
Dim SHisDate, SPosDate  As Variant
Dim AnnualSalary As Double, Compa!, SalaryGrade$
Dim xPosEarly
Dim xSH_PREMIUM, xSH_TOTAL, xSH_VGROUP, xSH_VSTEP
Dim xSHID 'George added Mar 9,2006 #9965

On Error GoTo UpRel_Err

JobCode$ = rsEPos("JH_JOB") 'xJob

If IsNumeric(rsEPos("JH_ID")) Then lngJobID& = rsEPos("JH_ID") Else lngJobID& = 0

PositionStartDat = rsEPos("JH_SDATE")
If Not IsNull(rsEPos("JH_WHRS")) Then HoursPerWeek! = rsEPos("JH_WHRS")
If Not IsNull(rsEPos("JH_JREASON")) Then JobReason$ = rsEPos("JH_JREASON")


SQLQ = "SELECT * FROM HR_SALARY_HISTORY"
SQLQ = SQLQ & " WHERE SH_EMPNBR = " & xEmpNo
SQLQ = SQLQ & " ORDER BY SH_EDATE DESC, SH_CURRENT " & IIf(glbSQL Or glbOracle, "DESC", "")
dynHRSALHIS.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If dynHRSALHIS.BOF And dynHRSALHIS.EOF Then
    'Msg = "No salary records found - New Employee?" & Chr(10)
    'Msg = Msg & "Please review and update this Employee's" & Chr(10)
    'Msg = Msg & "salary."
    'MsgBox Msg
    dynHRSALHIS.Close
    Exit Sub
End If

SHisDate = CVDate(dynHRSALHIS("SH_EDATE"))
If IsNull(dynHRSALHIS("SH_SDATE")) Then 'Ticket #24074 Franks 07/17/2013
    SPosDate = Date
Else
    SPosDate = CVDate(dynHRSALHIS("SH_SDATE"))
End If

''Ticket #24096 - Since now the Start Date can be same as previous record's start date, I now have to check if Job Codes of Prv and New
''Job is same then only update the Start Date. Also when changing an existing position only.
'If Not fglbNew Then
'    'xPosEarly = DateDiff("d", PositionStartDat, SHisDate) > 0
'    xPosEarly = DateDiff("d", PositionStartDat, SHisDate) > 0 And dynHRSALHIS("SH_JOB") = JobCode$
'    If xPosEarly Then
'        If fgtxtStartDate = SHisDate And dynHRSALHIS("SH_JOB") = JobCode$ Then
'            dynHRSALHIS("SH_SDATE") = CVDate(PositionStartDat)
'            dynHRSALHIS.Update
'            Exit Sub
'        End If
'    End If
'End If

xPosEarly = DateDiff("d", PositionStartDat, SHisDate) > 0 And dynHRSALHIS("SH_JOB") = JobCode$

'Ticket #24064 - Jerry said to disable this logic of adding a new salary record when Position and/or Start Date changes.
If True Then ' fglbNew Then
    dynHRSALHIS("SH_CURRENT") = False
    dynHRSALHIS.Update
    xSHID = dynHRSALHIS("SH_ID")
    'George added Mar 9,2006 #9965
    'If glbCompSerial = "S/N - 2259W" Or glbGP Then
    '    Call Salary_Integration(xEmpNo, , False, False, xSHID)
    'End If
    'George added Mar 9,2006 #9965

    If Not IsNull(dynHRSALHIS.Fields("SH_SALARY")) Then SH_SALARY@ = dynHRSALHIS.Fields("SH_SALARY")
    If Not IsNull(dynHRSALHIS.Fields("SH_SALCD")) Then SH_SALCD$ = dynHRSALHIS.Fields("SH_SALCD")
    If Not IsNull(dynHRSALHIS.Fields("SH_PAYP")) Then SH_PAYP$ = dynHRSALHIS.Fields("SH_PAYP")
    If Not IsNull(dynHRSALHIS.Fields("SH_NEXTDAT")) Then SH_NEXTDAT = dynHRSALHIS.Fields("SH_NEXTDAT")
    If glbWFC Then
        xSH_FISCALYEAR = "": xSH_SECTION = "": xSH_MARKETLINE = "": xSH_BAND = ""
        xPayrollID = ""
        xSH_CURRENCYINDI = ""
        If Not IsNull(dynHRSALHIS.Fields("SH_FISCALYEAR")) Then xSH_FISCALYEAR = dynHRSALHIS.Fields("SH_FISCALYEAR")
        If Not IsNull(dynHRSALHIS.Fields("SH_SECTION")) Then xSH_SECTION = dynHRSALHIS.Fields("SH_SECTION")
        If Not IsNull(dynHRSALHIS.Fields("SH_MARKETLINE")) Then xSH_MARKETLINE = dynHRSALHIS.Fields("SH_MARKETLINE")
        'If Len(clpCode(6).Text) > 0 Then xSH_BAND = clpCode(6).Text
        If Not IsNull(dynHRSALHIS.Fields("SH_BAND")) Then xSH_BAND = dynHRSALHIS.Fields("SH_BAND")
        If Not IsNull(dynHRSALHIS.Fields("SH_PAYROLL_ID")) Then xPayrollID = dynHRSALHIS.Fields("SH_PAYROLL_ID")
        If Not IsNull(dynHRSALHIS.Fields("SH_CURRENCYINDI")) Then xSH_CURRENCYINDI = dynHRSALHIS.Fields("SH_CURRENCYINDI")
    End If
    If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
        If Not IsNull(dynHRSALHIS.Fields("SH_PREMIUM")) Then xSH_PREMIUM = dynHRSALHIS.Fields("SH_PREMIUM")
        If Not IsNull(dynHRSALHIS.Fields("SH_TOTAL")) Then xSH_TOTAL = dynHRSALHIS.Fields("SH_TOTAL")
        If Not IsNull(dynHRSALHIS.Fields("SH_VGROUP")) Then xSH_VGROUP = dynHRSALHIS.Fields("SH_VGROUP")
        If Not IsNull(dynHRSALHIS.Fields("SH_VSTEP")) Then xSH_VSTEP = dynHRSALHIS.Fields("SH_VSTEP")
    End If

    'SET COMPA RATIO
    '================
    'Days and Months added by Bryan 30/Sep/05 Ticket#9354
    If JobSnap_Salary_Code$ = "A" Then
        If SH_SALCD$ = "H" Then
            AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 52
        ElseIf SH_SALCD$ = "M" Then
            AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 12
        ElseIf SH_SALCD$ = "D" Then
            If GetLeapYear(Year(Date)) Then
                AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 366
            Else
                AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 265
            End If
        Else
            AnnualSalary = SH_SALARY@
        End If
    ElseIf JobSnap_Salary_Code$ = "H" Then
        If SH_SALCD$ = "A" Then
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 52
        ElseIf SH_SALCD$ = "M" Then
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 12
        ElseIf SH_SALCD$ = "D" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 366
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 365
            End If
        Else
            AnnualSalary = SH_SALARY@
        End If
    ElseIf JobSnap_Salary_Code$ = "M" Then
        If SH_SALCD$ = "A" Then
            AnnualSalary = (SH_SALARY@ / HoursPerWeek!) / 12
        ElseIf SH_SALCD$ = "M" Then
            AnnualSalary = SH_SALARY@
        ElseIf SH_SALCD$ = "D" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * 366) / 12
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * 365) / 12
            End If
        Else
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * HoursPerWeek!) * 52 / 12
        End If
    ElseIf JobSnap_Salary_Code$ = "D" Then
        If SH_SALCD$ = "H" Then
            If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = (SH_SALARY@ * HoursPerWeek!) / 52
        ElseIf SH_SALCD$ = "M" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ * 12 / 366
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ * 12 / 365
            End If
        ElseIf SH_SALCD$ = "A" Then
            If GetLeapYear(Year(Date)) Then
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ / 366
            Else
                If HoursPerWeek! = 0 Then AnnualSalary = 0 Else AnnualSalary = SH_SALARY@ / 365
            End If
        Else
            AnnualSalary = SH_SALARY@
        End If
    End If

    ' set COMPA RATIO
    If glbWFC Then 'Ticket #25054 Franks 02/12/2014
        Compa! = Get_WFC_COMPA_FromMaster(glbUNION, JobCode$, SH_SALARY@, dynHRSALHIS.Fields("SH_SECTION"), dynHRSALHIS.Fields("SH_MARKETLINE"), dynHRSALHIS.Fields("SH_FISCALYEAR"))
    Else
        'If JobSnap_PayScale(JobSnap_MidPoint!) <> 0 And AnnualSalary <> 0 Then
        '    Compa! = (AnnualSalary / JobSnap_PayScale(JobSnap_MidPoint!)) * 100
        'Else
            Compa! = 0
        'End If
    End If
    
    If Compa! > 999.99 Then
        Compa! = 999.99
    End If

    'Determine Pay Scale individual fits into
    '==========================================
    SalaryGrade$ = "00"
    '''Ticket #22682 - Release 8.0: Increased the grid steps from 11 to 15 -> 20
    '''For x! = 1 To 11
    '''For X! = 1 To 15
    ''For x! = 1 To 20
    ''    If AnnualSalary >= JobSnap_PayScale(x) And JobSnap_PayScale(x) > 0 Then
    ''      cX$ = CStr(x)
    ''      If x! <= 9 Then cX$ = "0" & cX$
    ''      SalaryGrade$ = cX$
    ''    End If
    ''Next x!

    'NOW UPDATE SALARY HISTORY TABLE  - only if new record do we add record
    '================================
    If DateDiff("d", PositionStartDat, SHisDate) > 0 And glbSetPos Then GoTo SkipSal_Change

        If Not xPosEarly Then dynHRSALHIS.AddNew

        dynHRSALHIS("SH_COMPNO") = "001" 'SH_COMPNO%
        dynHRSALHIS("SH_EMPNBR") = xEmpNo
        dynHRSALHIS("SH_CURRENT") = True
        dynHRSALHIS("SH_SDATE") = CVDate(PositionStartDat)
        dynHRSALHIS("SH_EDATE") = IIf(xPosEarly, SHisDate, CVDate(PositionStartDat))
        dynHRSALHIS("SH_TRANSDATE") = Format(Now, "SHORT DATE")
        dynHRSALHIS("SH_SALARY") = SH_SALARY@
        dynHRSALHIS("SH_SALCD") = SH_SALCD$
        dynHRSALHIS("SH_JOB") = JobCode$
        'dynHRSALHIS("SH_GRID") = clpGrid.Text
        If Len(xPayrollID) > 0 Then dynHRSALHIS("SH_PAYROLL_ID") = xPayrollID
        'lngJobID&
        dynHRSALHIS("SH_JOB_ID") = lngJobID&
        dynHRSALHIS("SH_PAYP_TABLE") = "SDPP"
        dynHRSALHIS("SH_PAYP") = SH_PAYP$
        If IsDate(SH_NEXTDAT) Then
            If CVDate(SH_NEXTDAT) > IIf(xPosEarly, SHisDate, CVDate(PositionStartDat)) Then
                dynHRSALHIS("SH_NEXTDAT") = SH_NEXTDAT
            End If
        End If
        dynHRSALHIS("SH_WHRS") = HoursPerWeek!
        dynHRSALHIS("SH_SREAS_TABLE") = "SDRC"
        dynHRSALHIS("SH_SREAS1") = JobReason$     ' reason code
        dynHRSALHIS("SH_COMPA") = Round(Compa!, 2)
        dynHRSALHIS("SH_GRADE") = Format(SalaryGrade$, "00")
        dynHRSALHIS("SH_LDATE") = Date
        dynHRSALHIS("SH_LTIME") = Time$
        dynHRSALHIS("SH_LUSER") = glbUserID
        If glbWFC Then
            If Len(xSH_FISCALYEAR) > 0 Then
                dynHRSALHIS("SH_FISCALYEAR") = xSH_FISCALYEAR
            End If
            If Len(xSH_SECTION) > 0 Then
                dynHRSALHIS("SH_SECTION") = xSH_SECTION
            End If
            If Len(xSH_MARKETLINE) > 0 Then
                dynHRSALHIS("SH_MARKETLINE") = xSH_MARKETLINE
            End If
            If Len(xSH_BAND) > 0 Then
                dynHRSALHIS("SH_BAND") = xSH_BAND
            End If
            If Len(xSH_CURRENCYINDI) > 0 Then
                dynHRSALHIS("SH_CURRENCYINDI") = xSH_CURRENCYINDI
            End If
        End If
        If glbCompSerial = "S/N - 2373W" Then 'District Muskoka
            If Len(xSH_PREMIUM) > 0 Then
                dynHRSALHIS("SH_PREMIUM") = xSH_PREMIUM
            End If
            If Len(xSH_TOTAL) > 0 Then
                dynHRSALHIS("SH_TOTAL") = xSH_TOTAL
            End If
            If Len(xSH_VGROUP) > 0 Then
                dynHRSALHIS("SH_VGROUP") = xSH_VGROUP
            End If
            If Len(xSH_VSTEP) > 0 Then
                dynHRSALHIS("SH_VSTEP") = xSH_VSTEP
            End If
        End If
        dynHRSALHIS.Update

SkipSal_Change:
        xSHID = dynHRSALHIS("SH_ID")

        'Ticket #27056 - Update Audit table with this new Salary record
        If xIsUptAuditBen = "Y" Then
            If Not xPosEarly Then
                Call AUDITSALY_PUB(dynHRSALHIS("SH_EMPNBR"), dynHRSALHIS("SH_SALARY"), dynHRSALHIS("SH_PAYP"), dynHRSALHIS("SH_JOB"), dynHRSALHIS("SH_GRID"), dynHRSALHIS("SH_PAYROLL_ID"), dynHRSALHIS("SH_SALCD"), dynHRSALHIS("SH_WHRS"), dynHRSALHIS("SH_EDATE"), IIf(Not IsDate(dynHRSALHIS("SH_NEXTDAT")), Null, dynHRSALHIS("SH_NEXTDAT")), dynHRSALHIS("SH_SREAS1"))
            End If
        End If

    dynHRSALHIS.Close
    
    If xIsUptAuditBen = "Y" Then
        Call updBenefitForSalDEPN(xEmpNo)
    End If

    '''City of Niagara Falls - Ticket #15542
    ''If glbVadim And glbCompSerial = "S/N - 2276W" Then
    ''    'Add the salary record in Vadim's HR_EMP_HIST table storing the history of Rate changes
    ''    Call Update_VadimDB_HR_EMP_HISTORY(txtPayrollID, IIf(xPosEarly, SHisDate, CVDate(PositionStartDat)), "", Val(SalaryGrade$), JobCode$, "A")
    ''End If
    ''
    '''George added Mar 9,2006 #9965
    ''If glbCompSerial = "S/N - 2259W" Or glbGP Then 'Or (glbWFC And glbPlantCode = "GREN") Then
    ''    Call Salary_Integration(xEmpNo, , False, IIf(xPosEarly, False, True), xSHID)
    ''End If
    '''George added Mar 9,2006 #9965

End If


Exit Sub

UpRel_Err:
If Err = 3021 Then
    Exit Sub
End If
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err

'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SAL HISTORY", "HRSAL/PERF", "INSERT")
'Call RollBack '26July99 js

'End Sub

End Sub

Public Function WFC_AUDITBENF(xEmpNo, xlocNewRec As Boolean, rslBen As ADODB.Recordset, OBCode, OPPAMT, OBAMT, OEDate, OCOVER, OTCOST)
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
Dim ACTX
Dim NBCode, NPPAMT, NMTHCOMP, NMTHEMP, NBAMT, NPPE, NPCC, NMAXDOL, NEDate, NCOVER, NTCOST
Dim NBenGrp, NCeasedate
'''On Error GoTo AUDIT_ERR
WFC_AUDITBENF = False

If xlocNewRec Then
    ACTX = "A"
Else
    ACTX = "M"
End If

'If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
'Else
'    SQLQ = "SELECT ED_PT,ED_DIV FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
'    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
'End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDIV = ""
    Else
        xDIV = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDIV = ""
End If
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

xADD = False

NBCode = ""
NPPAMT = ""
NMTHCOMP = ""
NMTHEMP = ""
NBAMT = ""
NPPE = ""
NPCC = ""
NMAXDOL = ""
NEDate = ""
NCOVER = ""
NTCOST = ""
NBCode = rslBen("BF_BCODE")
NBenGrp = ""
NCeasedate = ""
If Not IsNull(rslBen("BF_PPAMT")) Then NPPAMT = rslBen("BF_PPAMT")
If Not IsNull(rslBen("BF_MTHCCOST")) Then NMTHCOMP = rslBen("BF_MTHCCOST")
If Not IsNull(rslBen("BF_MTHECOST")) Then NMTHEMP = rslBen("BF_MTHECOST")
If Not IsNull(rslBen("BF_AMT")) Then NBAMT = rslBen("BF_AMT")
If Not IsNull(rslBen("BF_PCC")) Then NPCC = rslBen("BF_PCC")
If Not IsNull(rslBen("BF_PCE")) Then NPPE = rslBen("BF_PCE")
If Not IsNull(rslBen("BF_MAXDOL")) Then NMAXDOL = rslBen("BF_MAXDOL")
If Not IsNull(rslBen("BF_EDATE")) Then NEDate = rslBen("BF_EDATE")
If Not IsNull(rslBen("BF_COVER")) Then NCOVER = rslBen("BF_COVER")
If Not IsNull(rslBen("BF_TCOST")) Then NTCOST = rslBen("BF_TCOST")
If Not IsNull(rslBen("BF_GROUP")) Then NBenGrp = rslBen("BF_GROUP")
If Not IsNull(rslBen("BF_CEASEDATE")) Then NCeasedate = rslBen("BF_CEASEDATE")

If OBCode <> NBCode Then GoTo MODUPD
'If OPPE <> NPPE Or OPCC <> NPCC Then GoTo MODUPD
'If OPPAMT <> NPPAMT Or OMAXDOL <> NMAXDOL Then GoTo MODUPD
If OPPAMT <> NPPAMT Then GoTo MODUPD
'If OMTHCOMP <> NMTHCOMP Or OMTHEMP <> NMTHEMP Then GoTo MODUPD
If OTCOST <> NTCOST Then GoTo MODUPD
If OBAMT <> NBAMT Then GoTo MODUPD
If OEDate <> NEDate Then GoTo MODUPD 'Ticket #23247 Franks 09/16/2013
If OBenGrp <> NBenGrp Then GoTo MODUPD 'Ticket #23247 Franks 09/16/2013

GoTo MODNOUPD

MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDIV

rsTA("AU_BCODE") = NBCode 'clpCode(1).Text
'If OMTHCOMP <> NMTHCOMP Then rsTA("AU_MTHCCOST") = NMTHCOMP
'If OMTHEMP <> NMTHEMP Then rsTA("AU_MTHECOST") = NMTHEMP
'If OTAXBEN <> txtTAXBEN Then rsTA("AU_TAXBEN") = txtTAXBEN
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'changed by raubrey 7/9/97 to make sure benefit code is written

  'If glbWFC Then 'Ticket #13772 Save the old Bcode here since wfc needs the old Bcode.
  '    If Len(OBCode) > 0 Then
  '        rsTA("AU_OLDLOC") = Left(OBCode, 10)
  '        If IsNumeric(OTOTAL) Then
  '            rsTA("AU_OLDWHRS") = OTOTAL
  '        End If
  '    End If
  'End If
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
If OCOVER <> NCOVER Then rsTA("AU_COVER") = NCOVER
If OTCOST <> NTCOST Then
    rsTA("AU_TCOST") = NTCOST
    rsTA("AU_MTHCCOST") = NMTHCOMP
    rsTA("AU_MTHECOST") = NMTHEMP
End If
'If OPremium <> lblAP Then rsTA("AU_PREMIUM") = lblAP
'If OPPE <> NPPE Then rsTA("AU_PCE") = NPPE
'If OPCC <> NPCC Then rsTA("AU_PCC") = NPCC
If OPPAMT <> NPPAMT Then
    rsTA("AU_PPAMT") = NPPAMT
    If IsNumeric(OPPAMT) Then rsTA("AU_OLDPPMT") = Val(OPPAMT)
End If
'If OMAXDOL <> NMAXDOL Then rsTA("AU_MAXDOL") = NMAXDOL
If OEDate <> NEDate Then
  If IsDate(NEDate) Then
      rsTA("AU_EDATE") = CVDate(NEDate)
  End If
End If
'If OPER <> txtPer Then rsTA("AU_PER") = txtPer
If OBAMT <> NBAMT Then rsTA("AU_BAMT") = NBAMT
'If OUNITCOST <> medUnitCost Then rsTA("AU_UNITCOST") = IIf(medUnitCost = "", 0, medUnitCost)

'If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
'Else
'    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
'    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
'End If
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo

rsTA("AU_LDATE") = Date
If IsDate(NEDate) Then 'if benefit effe date is future date, use it as LDATE
    If CVDate(NEDate) > CVDate(Date) Then
        rsTA("AU_LDATE") = CVDate(NEDate)
    End If
End If
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update
rsTA.Close

MODNOUPD:
WFC_AUDITBENF = True
Exit Function
AUDIT_ERR:

End Function

Public Function IsWFCUSBenEmp(xEmpNo)
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    
    If Not glbWFC Then IsWFCUSBenEmp = retVal: Exit Function
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    'SQLQ = SQLQ & "AND LEN(ED_VADIM1 )>0 " 'ED_VADIM1 - NGS Sub Group
    'comment out the code above since it caused a problem on Transfer In if ther is no ED_VADIM1 in the previous field
    SQLQ = SQLQ & "AND NOT (ED_EMP = 'CB') "
    SQLQ = SQLQ & "AND ED_WORKCOUNTRY = 'U.S.A.' " 'USA employee only
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        retVal = True
    End If
    IsWFCUSBenEmp = retVal
End Function
Public Function ESS_To_Tracker_EMP() 'Ticket #28373 Franks 03/29/2016
Dim ATConnectionString
Dim cnAT As New ADODB.Connection
Dim rsAT As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsIHR As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpnber, xDIV, xSection, xDOA, xReason
Dim updPayrollID
Dim xPaidType
Dim Mag, a%
Dim xUptFlag As Boolean
Dim xEmpCnt As Long
Dim retVal As Integer
    
    retVal = 0
    
    SQLQ = "SELECT * FROM HREMP WHERE (ED_EMPNBR IN (SELECT DISTINCT AU_EMPNBR FROM HRAUDIT WHERE LEFT(AU_SOURCE,3) = 'ESS' AND AU_INTUPLOAD = 'N')) "
    SQLQ = SQLQ & "OR (ED_EMPNBR IN (SELECT DISTINCT AU_EMPNBR FROM HRAUDIT2 WHERE LEFT(AU_SOURCE,3) = 'ESS' AND AU_INTUPLOAD = 'N')) "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xEmpCnt = 0
    If rsEmp.EOF Then
        ESS_To_Tracker_EMP = retVal
        Exit Function
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).Caption = "ESS Changes To Tracker..." ' "Updating ESS Changes To Tracker(Employees)..."
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
    
    '--- setup Sick Entitlement From Date and To Date if they are is blank, otherwise these caused error on ESS Entitlement Overview page
    SQLQ = "UPDATE HREMP SET ED_EFDATES = ED_EFDATE WHERE ED_EFDATES IS NULL"
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HREMP SET ED_ETDATES = ED_ETDATE WHERE ED_ETDATES IS NULL"
    gdbAdoIhr001.Execute SQLQ
    
    If Not rsEmp.EOF Then
        xEmpCnt = rsEmp.RecordCount
    End If
    Do While Not rsEmp.EOF
        xEmpnber = rsEmp("ED_EMPNBR")
            
        If glbWFC Then
            If Not IsNull(rsEmp("ED_BONUSDEPT")) Then
                If rsEmp("ED_BONUSDEPT") = "000000" Then
                    GoTo Next_Rec
                End If
            End If
        End If
        
        Call ESS_To_Tracker_By_Emp(rsEmp)
        
Next_Rec:
        SQLQ = "UPDATE HRAUDIT SET AU_INTUPLOAD = 'Y' WHERE AU_EMPNBR = " & xEmpnber & " AND AU_INTUPLOAD = 'N'"
        gdbAdoIhr001.Execute SQLQ
        SQLQ = "UPDATE HRAUDIT2 SET AU_INTUPLOAD = 'Y' WHERE AU_EMPNBR = " & xEmpnber & " AND AU_INTUPLOAD = 'N'"
        gdbAdoIhr001.Execute SQLQ
        rsEmp.MoveNext
        retVal = retVal + 1
    Loop
    rsEmp.Close
    
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
    
    Screen.MousePointer = DEFAULT
    
    ESS_To_Tracker_EMP = retVal
    
End Function

Public Function ESS_To_Tracker_ATT() 'Ticket #28373 Franks 03/29/2016
Dim ATConnectionString
Dim cnAT As New ADODB.Connection
Dim rsAT As New ADODB.Recordset
Dim rsAtt As New ADODB.Recordset
Dim rsIHR As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpnber, xDIV, xSection, xDOA, xReason
Dim updPayrollID
Dim xPaidType
Dim Mag, a%
Dim xUptFlag As Boolean
Dim xEmpCnt As Long
Dim retVal As Integer
    
    retVal = 0
    
    'SQLQ = "SELECT * FROM HREMP WHERE (ED_EMPNBR IN (SELECT DISTINCT AU_EMPNBR FROM HRAUDIT WHERE LEFT(AU_SOURCE,3) = 'ESS' AND AU_INTUPLOAD = 'N')) "
    'SQLQ = SQLQ & "OR (ED_EMPNBR IN (SELECT DISTINCT AU_EMPNBR FROM HRAUDIT2 WHERE LEFT(AU_SOURCE,3) = 'ESS' AND AU_INTUPLOAD = 'N')) "
    'RSATT.Open SQLQ, gdbAdoIhr001, adOpenStatic
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE LEFT(AD_SOURCE,3) = 'ESS' AND AD_UPLOAD = 'N'"
    rsAtt.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    xEmpCnt = 0
    If rsAtt.EOF Then
        ESS_To_Tracker_ATT = retVal
        Exit Function
    End If
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).Caption = "ESS Changes To Tracker..." '"Updating ESS Changes To Tracker(Attendances)..."
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
    
   
    If Not rsAtt.EOF Then
        xEmpCnt = rsAtt.RecordCount
    End If
    Do While Not rsAtt.EOF
        xEmpnber = rsAtt("AD_EMPNBR")
            
        ''If glbWFC Then
        ''    If Not IsNull(RSATT("ED_BONUSDEPT")) Then
        ''        If RSATT("ED_BONUSDEPT") = "000000" Then
        ''            GoTo Next_Rec
        ''        End If
        ''    End If
        ''End If
        
        ''Call ESS_To_Tracker_By_Emp(RSATT)
        Call WFC_Attend_To_AT(xEmpnber, "M", rsAtt("AD_DOA"), rsAtt("AD_REASON"), rsAtt("AD_ATT_ID"), , , "Y")
        
Next_Rec:
        rsAtt.MoveNext
        retVal = retVal + 1
    Loop
    rsAtt.Close
    
    'SQLQ = "UPDATE HR_ATTENDANCE SET AD_UPLOAD = 'Y' WHERE AD_EMPNBR = " & xEmpnber & " AND LEFT(AD_SOURCE,3) = 'ESS' AND AD_UPLOAD = 'N'"
    SQLQ = "UPDATE HR_ATTENDANCE SET AD_UPLOAD = 'Y' WHERE LEFT(AD_SOURCE,3) = 'ESS' AND AD_UPLOAD = 'N'"
    gdbAdoIhr001.Execute SQLQ
        
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(1).Caption = " "
    MDIMain.panHelp(2).Caption = " "
    MDIMain.panHelp(3).Caption = " "
    
    Screen.MousePointer = DEFAULT
    
    ESS_To_Tracker_ATT = retVal
End Function

Public Sub ESS_To_Tracker_By_Emp(rsEmp As ADODB.Recordset)  'Ticket #28373 Franks 03/29/2016
'Same logic as "Sub AT_Employee_Master_Update" in IHRIntegration
Dim cnAT As New ADODB.Connection
Dim rsAT As New ADODB.Recordset
'Dim rsEmp As New ADODB.Recordset
Dim rsIHR As New ADODB.Recordset
Dim xEmpnbr, updEMPID
Dim ATConnectionString
Dim SQLQ As String
Dim xDIV, xSection, xDOA, xReason
Dim updPayrollID As String, OldPayrollID As String, TemTableName As String
Dim xFunction

On Error Resume Next
    xFunction = "Employee Master"
    If IsNull(rsEmp("ED_SECTION")) Then xSection = "" Else xSection = rsEmp("ED_SECTION")
    If Not isTransferAT("Advanced Tracker", xFunction, xSection) Then Exit Sub

    ATConnectionString = OtherDatabaseInte("Advanced Tracker", xSection)
    If ATConnectionString = "" Then Exit Sub

    xTable = "[" & TableNamePrefix(xSection) & "-etp.etp.employee]"
    cnAT.CursorLocation = adUseClient
    cnAT.Open ATConnectionString
    
    xEmpnbr = rsEmp("ED_EMPNBR")
    updEMPID = rsEmp("ED_EMPNBR")
    If glbWFC Then 'Woodbridge North Carolina
        'Use Payroll ID instead of Employee # because it contains Division which AT does not want that
        'updPayrollID = GetEmpData(updEMPID, "ED_PAYROLL_ID", , Term_SEQ)
        If IsNull(rsEmp("ED_PAYROLL_ID")) Then updPayrollID = "" Else updPayrollID = rsEmp("ED_PAYROLL_ID")
        rsAT.Open "SELECT * FROM " & xTable & " as employee WHERE emp_code='" & updPayrollID & "'", cnAT, adOpenDynamic, adLockOptimistic
    Else
        rsAT.Open "SELECT * FROM " & xTable & " as employee WHERE emp_code='" & updEMPID & "'", cnAT, adOpenDynamic, adLockOptimistic
    End If
    
    '------------- update begin ------------------------------------------------
    If Not rsAT.EOF Then 'only for update, no new record

        If Not IsNull(rsEmp!ED_BUSNBR) Then
            rsAT!altphone = Format(Left(rsEmp!ED_BUSNBR, 10), "######-####")
        End If
        rsAT!birthdate = rsEmp!ED_DOB
        If Not IsNull(rsEmp!ED_CELLPHONE) Then
            rsAT!cellphone = Format(Left(rsEmp!ED_CELLPHONE, 10), "######-####")
        End If
        
        If glbCompSerial = "S/N - 2296W" Then 'SERIAL_EssexLib
            rsAT!defaulttimecalc = Left(rsEmp!ED_USER_TEXT1, 10)
        End If
        rsAT!doctorname = Left(rsEmp!ED_EDOCTOR, 20)
        rsAT!doctornum = Left(rsEmp!ED_EDPNBR, 15)
        rsAT!email = Left(rsEmp!ED_EMAIL, 40)
        rsAT!emercontactname = Left(rsEmp!ED_ECONT, 20)
        rsAT!emercontactnum1 = Left(rsEmp!ED_ENBR, 15)
        rsAT!emercontactnum2 = Left(rsEmp!ED_EP2NBR, 15)
        rsAT!emercontactname2 = Left(rsEmp!ED_ECONT2, 20)
        rsAT!emercontactnum21 = Left(rsEmp!ED_ENBR2, 15)
        rsAT!emercontactnum22 = Left(rsEmp!ED_EP2NBR2, 15)
        rsAT!emercontact1relation = Left(rsEmp!ED_RELATE, 15)
        rsAT!emercontact2relation = Left(rsEmp!ED_RELATE2, 15)
        
        If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2390W" Or glbCompSerial = "S/N - 2296W" Or glbCompSerial = "S/N - 2301W" Then
        'Hemu - Ticket #15012 - St. John's - Do not pass
        'Ticket #25376 - Community Living Access Support Services - Do not pass Apartment Num
        Else
            rsAT!emp_apartnum = Left(rsEmp!ED_ADDR2, 40)
        End If
        If glbCompSerial = "S/N - 2394W" Then  'Hemu - Ticket #15012 - St. John's - Do not pass
        Else
            rsAT!emp_city = Left(rsEmp!ED_CITY, 20)
        End If
        rsAT!emp_firstname = Left(rsEmp!ED_FNAME, 20)
        rsAT!emp_lastname = Left(rsEmp!ED_SURNAME, 20)
        If glbCompSerial = "S/N - 2296W" Then 'SERIAL_EssexLib
            'Ticket #18789 Franks 05/04/2011
        Else
            rsAT!emp_midname = Left(rsEmp!ED_MIDNAME, 20)
        End If
        
        If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2390W" Then  'Hemu - Ticket #15012 - St. John's - Do not pass
        Else
            rsAT!emp_postal_zip = Left(rsEmp!ED_PCODE, 20)
            rsAT!emp_prov_state = Left(rsEmp!ED_PROV, 2)
        End If
    
        If glbWFC Then 'Woodbridge North Carolina
            'Ticket #23948 Frank 06/24/2013 - begin
            'rsAT!emp_status = Left(rsEmp!ED_PT, 6)
            xWFCStatusSecurity = getWFCStatusSecurity(xSection)
            If xWFCStatusSecurity = "EDEMP" Then
                rsAT!emp_status = Left(rsEmp!ED_EMP, 6)
            Else
                rsAT!emp_status = Left(rsEmp!ED_PT, 6)
            End If
            'Ticket #23948 Frank 06/24/2013 -  end
        'ElseIf glbCompSerial = "S/N - 2394W" Then 'St. John's Ticket #14739
        '    If IfLoaCode(rsEmp!ED_EMP) Then
        '        rsAT!emp_status = Left(rsEmp!ED_EMP, 6)
        '    Else
        '        rsAT!emp_status = Left(rsEmp!ED_PT, 6)
        '    End If
        ElseIf glbCompSerial = "S/N - 2379W" Then 'Town of LaSalle Ticket #14534
            rsAT!emp_status = GetEmpStatus4LaSalle(rsEmp!ED_EMP, rsEmp!ed_emptype)
        ElseIf glbCompSerial = "S/N - 2296W" Then 'SERIAL_EssexLib
            rsAT!emp_status = Left(rsEmp!ED_SALDIST, 6)
        ElseIf glbCompSerial = "S/N - 2335W" Then 'Mitchell Plastic
            'rsAT!emp_status = Left(rsEmp!ED_PT, 6)
            'Ticket #21649 Franks 03/02/2012
            rsAT!emp_status = Left(rsEmp!ED_ORG, 6)
        ElseIf glbCompSerial = "S/N - 2301W" Then  'Ticket #25376 - Community Living Access Support Services
            'rsAT!emp_status = Left(rsEmp!ED_EMP & "-" & rsEmp!ED_PT, 6)
            rsAT!emp_status = Left(rsEmp!ED_EMP, 6)
        Else
            rsAT!emp_status = Left(rsEmp!ED_EMP, 6)
        End If
        If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2390W" Then  'Hemu - Ticket #15012 - St. John's - Do not pass
        Else
            rsAT!emp_street = Left(rsEmp!ED_ADDR1, 40)
        End If
        rsAT!emp_name = Left(rsEmp!ED_SURNAME & ", " & rsEmp!ED_FNAME, 35)
    
        ''If glbCompSerial = "S/N - 2355W" Then 'FOR LAMBTON
        ''    rsAT!emp_shift = Left(rsJOB!JH_SHIFT, 6)
        ''    If IsNull(rsAT!emp_shift) Then rsAT!emp_shift = "DEF"
        ''ElseIf glbCompSerial = "S/N - 2390W" Then 'Collectcorp
        ''    rsAT!emp_shift = "OPEN"
        ''ElseIf glbCompSerial = "S/N - 2296W" Then 'SERIAL_EssexLib
        ''    rsAT!emp_shift = "NOTSCH"
        ''Else
        ''    rsAT!emp_shift = Left(rsJOB!JH_SHIFT, 6)
        ''    If glbCompSerial <> "S/N - 2282W" Then 'Woodbridge North Carolina
        ''        If IsNull(rsAT!emp_shift) Then rsAT!emp_shift = "2"
        ''    End If
        ''End If
        If glbCompSerial = "S/N - 2282W" Then 'Woodbridge North Carolina
            'Ticket #11945, comment out the DELR
            'If glbMultiPayCode = "DELR" Then 'Del Rio
            '    If Not IsNull(rsEMP!ED_DEPTNO) Then
            '        rsAT!emp_dept = Mid(rsEMP!ED_DEPTNO, 5, Len(rsEMP!ED_DEPTNO) - 4)
            '    End If
            'Else
                'Both AT and IHR are text, leave the 0's - Bryan 17/Mar/2006
                If IsNumeric(Left(rsEmp!ED_DEPTNO, 6)) Then
                    rsAT!emp_dept = Format(Val(Left(rsEmp!ED_DEPTNO, 6)), "000000")
                Else
                    rsAT!emp_dept = Left(rsEmp!ED_DEPTNO, 6)
                End If
            'End If
        'ElseIf glbCompSerial = "S/N - 2394W" Then 'St. John's Rehab #14796
        '    rsAT!emp_dept = GetShortGLNO(rsEmp!ED_GLNO)
        Else
            rsAT!emp_dept = Left(rsEmp!ED_DEPTNO, 6)
        End If
        If glbCompSerial = "S/N - 2381W" Then 'The Elliot Community - Ticket #13285
            rsAT!emp_badge = Left(rsEmp!ED_EMPNBR, 10)
        ElseIf glbCompSerial = "S/N - 2379W" Or glbCompSerial = "S/N - 2296W" Then 'SERIAL_EssexLib 'Town of LaSalle Ticket #14534
            rsAT!emp_badge = Left(rsEmp!ED_EMPNBR, 10)
        Else
            rsAT!emp_badge = Left(rsEmp!ED_BADGEID, 10)
        End If
        
        '''update Salary - begin
        ''Dim tmpSal
        ''If Format(rsSal!SH_SALCD, "@") <> "H" Then
        ''    If Val(Format(rsJOB!JH_WHRS, "@")) = 0 Then
        ''        tmpSal = 0
        ''    Else
        ''        tmpSal = rsSal!SH_SALARY / rsJOB!JH_WHRS / 52
        ''    End If
        ''Else
        ''    tmpSal = rsSal!SH_SALARY
        ''End If
        ''If glbCompSerial = "S/N - 2282W" Then 'Woodbridge North Carolina 'Frank 03/18/2006
        ''    'If rsAT!emp_rate <> tmpSal Then UpdateRate Format(rsEMP!ED_PAYROLL_ID, "000000"), tmpSal, rsSal!SH_EDATE, cnAT
        ''    If rsAT!emp_rate <> tmpSal Then UpdateRate rsEmp!ED_PAYROLL_ID, tmpSal, rsSal!SH_EDATE, cnAT
        ''    If IsNull(rsAT!emp_rate) And Not IsNull(tmpSal) Then
        ''        'UpdateRate Format(rsEMP!ED_PAYROLL_ID, "000000"), tmpSal, rsSal!SH_EDATE, cnAT
        ''        UpdateRate rsEmp!ED_PAYROLL_ID, tmpSal, rsSal!SH_EDATE, cnAT
        ''    End If
        ''    rsAT!emp_rate = 0 'WFC Ticket #11825
        ''    'Ticket #12339
        ''    DelRate xEmpnbr, rsEmp!ED_PAYROLL_ID, cnAT
        ''ElseIf glbCompSerial = "S/N - 2394W" Then 'Ticket #14791
        ''    rsAT!emp_rate = 0
        ''    UpdateRate updEMPID, 0, rsSal!SH_EDATE, cnAT
        ''Else
        ''    If rsAT!Rate <> tmpSal Then UpdateRate updEMPID, tmpSal, rsSal!SH_EDATE, cnAT
        ''    If IsNull(rsAT!emp_rate) And Not IsNull(tmpSal) Then 'Ticket #14681
        ''         UpdateRate updEMPID, tmpSal, rsSal!SH_EDATE, cnAT
        ''    End If
        ''    rsAT!emp_rate = tmpSal
        ''    DelRate xEmpnbr, updEMPID, cnAT
        ''End If
        '''update Salary - end
        If glbCompSerial = "S/N - 2282W" Or glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2390W" Then
            'Woodbridge North Carolina, St. Johns, Collectcorp
            rsAT!emp_sdate = rsEmp!ED_DOH
        Else
            rsAT!emp_sdate = rsEmp!ED_USRDAT1
        End If
        If glbCompSerial = "S/N - 2282W" Or glbCompSerial = "S/N - 2390W" Then 'Woodbridge North Carolina
            rsAT!gender = IIf(rsEmp!ED_SEX = "F", "F", "M") 'Default to Male
        Else
            rsAT!gender = IIf(rsEmp!ED_SEX = "M", "Male", "Female")
        End If
        '?rsAT!healthcard = LEFT(ESEMP!ED_HEALTHCARD, 12)
        
        If glbCompSerial = "S/N - 2282W" Or glbCompSerial = "S/N - 2390W" Or glbCompSerial = "S/N - 2296W" Then 'SERIAL_EssexLib Then 'Woodbridge North Carolina, Collectcorp
            rsAT!holiday_pay = True
        Else
            If rsEmp!ED_PROV = "ON" Then rsAT!holiday_pay = True
        End If
        rsAT!hire_date = rsEmp!ED_DOH
        Dim strMarried As String
        If rsEmp("ED_MSTAT") = "S" Then strMarried = "Single"
        If rsEmp("ED_MSTAT") = "M" Then strMarried = "Married"
        If rsEmp("ED_MSTAT") = "F" Then strMarried = "Family"
        If rsEmp("ED_MSTAT") = "P" Then strMarried = "Parent(S)"
        If rsEmp("ED_MSTAT") = "D" Then strMarried = "Divorced"
        If rsEmp("ED_MSTAT") = "W" Then strMarried = "Widowed"
        If rsEmp("ED_MSTAT") = "C" Then strMarried = "Common-Law"
        If rsEmp("ED_MSTAT") = "R" Then strMarried = "Partner"
        If rsEmp("ED_MSTAT") = "X" Then strMarried = "Same-Sex"
        If rsEmp("ED_MSTAT") = "O" Then strMarried = "Other"
        If glbCompSerial = "S/N - 2282W" Then 'Woodbridge North Carolina
            rsAT!maritalstatus = UCase(strMarried)
        ElseIf glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2390W" Then    'Do not transfer ED_MSTAT for St. John's Ticket #14739
        Else
            rsAT!maritalstatus = strMarried
        End If
        If Not IsNull(rsEmp!ED_PAGENBR) Then
            If glbCompSerial = "S/N - 2390W" Or glbCompSerial = "S/N - 2296W" Then
            Else
            rsAT!pager = Format(Left(rsEmp!ED_PAGENBR, 10), "######-####")
            End If
        End If
        'If glbCompSerial = "S/N - 2355W" Or glbCompSerial = "S/N - 2390W" Then  'FOR LAMBTON
        'Else
        '    rsAT!payincreasedate = rsSal!SH_EDATE
        'End If
        If glbCompSerial = "S/N - 2394W" Or glbCompSerial = "S/N - 2390W" Then     'Hemu - Ticket #14752 - Do not transfer any phone #s. to AT
            'Ticket #20368 Franks 05/26/2011
            'Transfer “Home Telephone Number” and “Seniority Date” to Advanced Tracker if the employee’s “Clinical/Non Clin”
            '(Section code in info:HR)equals CLIN and if they are in Department A1, A2, A3 or A4.
            If glbCompSerial = "S/N - 2394W" Then
                If Not IsNull(rsEmp("ED_SECTION")) And Not IsNull(rsEmp("ED_DEPTNO")) Then
                    If rsEmp("ED_SECTION") = "CLIN" And (rsEmp("ED_DEPTNO") = "A1" Or rsEmp("ED_DEPTNO") = "A2" Or rsEmp("ED_DEPTNO") = "A3" Or rsEmp("ED_DEPTNO") = "A4") Then
                        rsAT!phone = Format(Left(rsEmp!ED_PHONE, 10), "######-####")
                    End If
                End If
            End If
        Else
            If Not IsNull(rsEmp!ED_PHONE) Then
                rsAT!phone = Format(Left(rsEmp!ED_PHONE, 10), "######-####")
            End If
        End If
        If glbCompSerial = "S/N - 2390W" Then
        Else
            If glbCompSerial = "S/N - 2282W" Then 'Woodbridge North Carolina
                rsAT!Sin_Ssn = Format(rsEmp!ED_SIN, "###-##-####")
            Else
                rsAT!Sin = Left(rsEmp!ED_SIN, 9)
            End If
        End If
    
    
        rsAT.Update
    End If
    '------------- update end ------------------------------------------------
    
    rsAT.Close
    cnAT.Close
    
Exit Sub

Err_Employee_Master_:
Resume Next

End Sub

Private Function getWFCStatusSecurity(xMultiPayCode) 'Ticket #23948 Frank 06/24/2013
Dim rsSetup As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = "EDPT"
    SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' "
    SQLQ = SQLQ & "AND PARA_CATEGORY='Advanced Tracker'  "
    SQLQ = SQLQ & "AND PARA_CATEGORY2='Integration Setup'  "
    SQLQ = SQLQ & "AND PARA_MULTIPAY_CODE = '" & xMultiPayCode & "' "
    SQLQ = SQLQ & "AND PARA_NAME = 'Use Emp Status' "
    If rsSetup.State <> 0 Then rsSetup.Close
    rsSetup.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If Not rsSetup.EOF Then
        If rsSetup("PARA_VALUE") = "True" Then
            retVal = "EDEMP"
        End If
    End If
    getWFCStatusSecurity = retVal
End Function
Private Function GetEmpStatus4LaSalle(xEmpStatus, xEmpType)
Dim xATEmpStatus As String
    xATEmpStatus = ""
    If xEmpStatus = "A" Then
        If Not IsNull(xEmpType) Then
            xATEmpStatus = xEmpStatus & Left(Trim(xEmpType), 1)
        Else
            xATEmpStatus = xEmpStatus
        End If
    Else
        xATEmpStatus = xEmpStatus
    End If
    GetEmpStatus4LaSalle = xATEmpStatus
End Function

Public Sub WFC_Attend_To_AT(xEmpNo, xType, xHRDate, xHRReason, xAttID, Optional xOldHRDate = "", Optional xOldReason = "", Optional IsFromESS = "N") 'Ticket #24124 Franks 07/24/2013
Dim ATConnectionString
Dim cnAT As New ADODB.Connection
Dim rsAT As New ADODB.Recordset
Dim rsEmp As New ADODB.Recordset
Dim rsIHR As New ADODB.Recordset
Dim rsMatrix As New ADODB.Recordset
Dim SQLQ As String
Dim xEmpnber, xDIV, xSection, xDOA, xReason, xPT
Dim xTabl
Dim shiftinfo
Dim updPayrollID
Dim xPaidType
Dim Mag, a%
Dim xUptFlag As Boolean
Dim xUseTimeBank As Boolean 'Ticket #27298 Franks 07/13/2015
Dim xJobCode
Dim xTotHrs

On Error GoTo errfalse
    
    If IsFromESS = "Y" Then 'Ticket #28373 Franks 03/30/2016
    Else
        If Not (glbAdv Or glbWFCFullRights) Then Exit Sub
    End If
    
    'find the matching ed_empnbr
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        rsEmp.Close
        Exit Sub
    End If
    xEmpnbr = rsEmp("ED_EMPNBR")
    xSection = rsEmp("ED_SECTION")
    If Not IsNull(rsEmp("ED_PT")) Then xPT = rsEmp("ED_PT") Else xPT = ""
    If rsEmp.State <> 0 Then rsEmp.Close
    
    'check if this function is turn on under Adv Integration Setup
    If Not isTransferAT("Advanced Tracker", "Export Attendance", xSection) Then Exit Sub
    
    'Lookup the Attendance Code Matrix to determine if the Attendance Reason is an unpaid or paid
    'reason code. If the code is not in the matrix, do not update Tracker.
    If IsNull(xHRReason) Then Exit Sub
    If Len(xHRReason) = 0 Then Exit Sub
    
    'Ticket #28427 Franks 04/19/2016 - begin
    'check the if there is record for the ED_PT(Category)
    xPaidType = ""
    If Len(xPT) > 0 Then
        SQLQ = "SELECT * FROM HRATT_MATRIX WHERE AM_REASON = '" & xHRReason & "' "
        SQLQ = SQLQ & "AND AM_SECTION = '" & xSection & "' "
        SQLQ = SQLQ & "AND ',' + AM_PT + ','  LIKE '%," & xPT & ",%'"
        rsMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xPaidType = ""
        If rsMatrix.EOF Then
        Else
            xPaidType = rsMatrix("AM_CODE_TYPE")
        End If
    End If
    'Ticket #28427 Franks 04/19/2016 - end

    If Len(xPaidType) = 0 Then
        SQLQ = "SELECT * FROM HRATT_MATRIX WHERE AM_REASON = '" & xHRReason & "' "
        SQLQ = SQLQ & "AND AM_SECTION = '" & xSection & "' "
        If rsMatrix.State <> 0 Then rsMatrix.Close
        rsMatrix.Open SQLQ, gdbAdoIhr001, adOpenStatic
        xPaidType = ""
        If rsMatrix.EOF Then
            rsMatrix.Close
            Exit Sub
        Else
            xPaidType = rsMatrix("AM_CODE_TYPE")
        End If
    End If
    
    'Ticket #27298 Franks 07/13/2015 - begin
    xUseTimeBank = False
    If Not rsMatrix.EOF Then
        If Not IsNull(rsMatrix("AM_ABSENT_HRS")) Then
            If rsMatrix("AM_ABSENT_HRS") Then
                xUseTimeBank = True
            End If
        End If
    End If
    'Ticket #27298 Franks 07/13/2015 - end
    rsMatrix.Close
    
    If xType = "D" Then 'delete begin
        If CVDate(xHRDate) < CVDate(Date) Then
            If IsFromESS = "Y" Then
                'No message pop up
            ElseIf IsFromESS = "U" Then  ''Ticket #28919 Franks 03/24/2017
                'No message pop up
            Else
                Msg = "Previously dated transactions will not post to Tracker. " '& Chr(10)
                Msg = Msg & "If changes need to be made in Tracker, they must be done manually." & Chr(10)
                MsgBox Msg
            End If
            Exit Sub
        Else 'CVDate(xHRDate) >= CVDate(Date)
            'find the matching Att in AT
            ATConnectionString = OtherDatabaseInte("Advanced Tracker", xSection)
            If ATConnectionString = "" Then Exit Sub
            
            If InStr(ATConnectionString, "SQLOLEDB") <> 0 Then
                xTable = "[" & TableNamePrefix(xSection) & "-etp.etp.timehistory]"
                cnAT.CursorLocation = adUseClient
                cnAT.Open ATConnectionString
            Else
                'xTable = "timehistory"
                'cnAT.Open ATConnectionString
                Exit Sub 'for Access - don't use it
            End If

            If glbWFC Then 'Woodbridge North Carolina
                'Use Payroll ID instead of Employee # because it contains Division which AT does not use it
                'updPayrollID = Format(getEmpData(UpdKeys(0), "ED_PAYROLL_ID"), "000000")
                updPayrollID = GetEmpData(xEmpnbr, "ED_PAYROLL_ID")
                SQLQ = "SELECT * FROM " & xTable & " WHERE thist_emp='" & updPayrollID & "'"
            Else
                SQLQ = "SELECT * FROM " & xTable & " WHERE thist_emp='" & xEmpnbr & "'"
            End If
            'SQLQ = SQLQ & " AND shift_date='" & UpdKeys(1) & "'"
            SQLQ = SQLQ & " AND shift_date = " & Date_SQL(xHRDate) & " "
            SQLQ = SQLQ & " AND thist_attend_code = '" & xHRReason & "' "
            rsAT.Open SQLQ, cnAT, adOpenKeyset, adLockOptimistic
            If Not rsAT.EOF Then 'found
                rsAT.Delete
            End If
            rsAT.Close
        End If
    End If 'delete end
    
    If xType = "M" Then '----------------

        'find the matching Att recrod
        SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & xEmpNo & " AND AD_ATT_ID = " & xAttID & " "
        rsIHR.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsIHR.EOF Then
            rsIHR.Close
            Exit Sub
        End If
        xDOA = rsIHR("AD_DOA")
        xReason = rsIHR("AD_REASON")
        
        'find the matching Att in AT
        ATConnectionString = OtherDatabaseInte("Advanced Tracker", xSection)
        If ATConnectionString = "" Then Exit Sub
        
        If InStr(ATConnectionString, "SQLOLEDB") <> 0 Then
            xTable = "[" & TableNamePrefix(xSection) & "-etp.etp.timehistory]"
            cnAT.CursorLocation = adUseClient
            cnAT.Open ATConnectionString
        Else
            'xTable = "timehistory"
            'cnAT.Open ATConnectionString
            Exit Sub 'for Access - don't use it
        End If
        
        'Ticket #24337 Franks 09/30/2013 - begin
        'o   User changes a Reason Code in info:HR and a new record is created in Tracker - should delete it and add new
        If Len(xOldReason) > 0 Then
            If Not xOldReason = xHRReason Then
                If glbWFC Then 'Woodbridge North Carolina
                    'Use Payroll ID instead of Employee # because it contains Division which AT does not use it
                    updPayrollID = GetEmpData(xEmpnbr, "ED_PAYROLL_ID")
                    SQLQ = "SELECT * FROM " & xTable & " WHERE thist_emp='" & updPayrollID & "'"
                Else
                    SQLQ = "SELECT * FROM " & xTable & " WHERE thist_emp='" & xEmpnbr & "'"
                End If
                'SQLQ = SQLQ & " AND shift_date='" & UpdKeys(1) & "'"
                SQLQ = SQLQ & " AND shift_date = " & Date_SQL(xHRDate) & " "
                SQLQ = SQLQ & " AND thist_attend_code = '" & xOldReason & "' "
                rsAT.Open SQLQ, cnAT, adOpenKeyset, adLockOptimistic
                If Not rsAT.EOF Then 'found
                    rsAT.Delete
                End If
                rsAT.Close
            End If
        End If
        'Ticket #24337 Franks 09/30/2013 - end
        
        xUptFlag = False
        If glbWFC Then 'Woodbridge North Carolina
            'Use Payroll ID instead of Employee # because it contains Division which AT does not use it
            'updPayrollID = Format(getEmpData(UpdKeys(0), "ED_PAYROLL_ID"), "000000")
            updPayrollID = GetEmpData(xEmpnbr, "ED_PAYROLL_ID")
            SQLQ = "SELECT * FROM " & xTable & " WHERE thist_emp='" & updPayrollID & "'"
        Else
            SQLQ = "SELECT * FROM " & xTable & " WHERE thist_emp='" & xEmpnbr & "'"
        End If
        If IsDate(xOldHRDate) Then 'Ticket #24155 Franks 07/30/2013
            SQLQ = SQLQ & " AND shift_date = " & Date_SQL(xOldHRDate) & " "
        Else
            SQLQ = SQLQ & " AND shift_date = " & Date_SQL(xDOA) & " "
        End If
        SQLQ = SQLQ & " AND thist_attend_code = '" & xReason & "' "
        rsAT.Open SQLQ, cnAT, adOpenKeyset, adLockOptimistic
        If Not rsAT.EOF Then 'found
            If CVDate(xDOA) < CVDate(Date) Then
                If IsFromESS = "Y" Then
                    xUptFlag = True
                ElseIf IsFromESS = "U" Then  ''Ticket #28919 Franks 03/24/2017
                    xUptFlag = True ''No message pop up
                Else
                    Msg = "Duplicate record exists in Tracker. " '& Chr(10)
                    Msg = Msg & "Do you want to update the Hours in Tracker to use info:HR's Hours?" & Chr(10)
                    a% = MsgBox(Msg, vbYesNo + vbQuestion, "Confirm")
                    If a% = vbNo Then
                        Exit Sub
                    Else
                        xUptFlag = True
                    End If
                End If
            Else 'CVDate(xDOA) >= CVDate(Date)
                xUptFlag = True 'update AT record
            End If
        Else 'not found
            If CVDate(xDOA) < CVDate(Date) Then
                If IsFromESS = "Y" Then
                    xUptFlag = True
                ElseIf IsFromESS = "U" Then  ''Ticket #28919 Franks 03/24/2017
                    xUptFlag = True ''No message pop up
                Else
                    Msg = "Do you want to update Tracker with this data? " '& Chr(10)
                    Msg = Msg & "If the Pay Period is closed or locked, you should not transfer this data to Tracker." & Chr(10)
                    a% = MsgBox(Msg, vbYesNo + vbQuestion, "Confirm")
                    If a% = vbNo Then
                        Exit Sub
                    Else
                        xUptFlag = True
                    End If
                End If
            Else 'CVDate(xDOA) >= CVDate(Date)
                xUptFlag = True 'create new record
            End If
        End If
        If xUptFlag Then 'update AT record
            'Ticket #28373 Franks 04/14/2016 -- begin
            If IsFromESS = "Y" Then
                xTotHrs = getTotHrsToESS(xEmpnbr, xDOA, xReason)
            Else
                xTotHrs = rsIHR!AD_HRS
            End If
            If xTotHrs = 0 Then 'DO NOT UPDATE TRACKER IF HOURS IS 0
                If IsFromESS = "Y" Then
                    'need 0 to update Tracker since it is total hours
                Else
                    Exit Sub
                End If
            End If
            'Ticket #28373 Franks 04/14/2016 -- end
            
            If rsAT.EOF Then
                rsAT.AddNew
                If IsFromESS = "Y" Then
                    rsAT!notes2 = "added by ESS on " & Now
                ElseIf IsFromESS = "U" Then  ''Ticket #28919 Franks 03/24/2017
                    rsAT!notes2 = "added by Info Att Muss Update on " & Now
                Else
                    rsAT!notes2 = "added by Info HR on " & Now
                End If
                If glbWFC Then 'Woodbridge North Carolina
                    If IsNumeric(GetEmpData(rsIHR!AD_EMPNBR, "ED_DEPTNO")) Then
                        rsAT!thist_dept = Val(GetEmpData(rsIHR!AD_EMPNBR, "ED_DEPTNO"))
                    Else
                        rsAT!thist_dept = GetEmpData(rsIHR!AD_EMPNBR, "ED_DEPTNO")
                    End If
                Else
                    rsAT!thist_dept = GetEmpData(rsIHR!AD_EMPNBR, "ED_DEPTNO")
                End If
                If glbWFC Then
                    'Ticket #27531 Franks 09/21/2015
                    '"   When entering attendance and the record goes automatically, make sure the position code is converted to the job code.
                    xJobCode = GetJHData(rsIHR!AD_EMPNBR, "JH_JOB", "")
                    xJobCode = getNewJobMasterCode(xJobCode)
                    rsAT!Position = Left(xJobCode, 6)
                Else
                    rsAT!Position = GetJHData(rsIHR!AD_EMPNBR, "JH_JOB", "")
                End If
            End If
            rsAT!thist_record_type = "AR"
            If glbCompSerial = "S/N - 2282W" Then 'Woodbridge North Carolina
                rsAT!thist_emp = updPayrollID
            Else
                rsAT!thist_emp = Left(rsIHR!AD_EMPNBR, 9)
            End If
            rsAT!SHIFT_DATE = rsIHR!AD_DOA
            rsAT!thist_date = rsIHR!AD_DOA
            rsAT!thist_attend_code = rsIHR!AD_REASON
            ''If rsIHR!AD_REASON = "VAC" Or Left(rsIHR!AD_REASON, 3) = "SIC" Then
            ''    rsAT!OTBankCode = rsIHR!AD_REASON
            ''    rsAT!thist_reference = "BA"
            ''ElseIf IsHourlyEnt(rsIHR!AD_EMPNBR, rsIHR!AD_REASON) Then
            ''    rsAT!OTBankCode = rsIHR!AD_REASON
            ''End If
            'Ticket #24155 Franks 07/30/2013 - leave the followin two field as blank as Jamie asked
            'rsAT!thist_reference = "BA"
            'rsAT!time_status = "B" 'new
            
            'Ticket #27298 Franks 07/13/2015
            If glbWFC Then
                If xUseTimeBank Then
                    rsAT!OTBankCode = rsIHR!AD_REASON
                    rsAT!thist_reference = "BA"
                    rsAT!time_status = "B"
                End If
            End If
            If rsIHR!AD_SHIFT & "" = "" Then
                rsAT!thist_shift = GetJHData(rsIHR!AD_EMPNBR, "JH_SHIFT", "")
            Else
                rsAT!thist_shift = rsIHR!AD_SHIFT & ""
            End If
            'If rsIHR!AD_REASON = "CT" Then
            '    rsAT!t15 = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
            'Else
            '    'If glbPaidHours Then 'Ticket #14739 ???
            '    '    rsAT!reg = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
            '    'End If
            'End If
            If xPaidType = "Paid" Then
                'rsAT!reg = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
                'rsAT!o_reg = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
                rsAT!reg = DateAdd("n", (xTotHrs * 60), "12:00 AM")
                rsAT!o_reg = DateAdd("n", (xTotHrs * 60), "12:00 AM")
            End If
            If glbCompSerial = "S/N - 2282W" Then 'Woodbridge North Carolina
                rsAT!Rate = getRateAT(updPayrollID, xSection)
            Else
                rsAT!Rate = getRateAT(rsIHR!AD_EMPNBR, xSection)
            End If
            rsAT!base_rate = rsAT!Rate
                    
            '''Ticket #24268 Franks 10/09/2013 - Total = Hours x rate
            ''rsAT!Total = DateAdd("n", (rsIHR!AD_HRS * 60 * rsAT!Rate), "12:00 AM")
            ''Ticket #24268 Franks 01/10/2014 - Total = Hours x rate
            'rsAT!Total = rsIHR!AD_HRS * rsAT!Rate
            rsAT!Total = xTotHrs * rsAT!Rate
            
            shiftinfo = Split(getShiftInfo(rsAT!thist_shift, xSection), "|")
            
            If glbMitchellPlastics And xSection = "UMW" Then
                'Ticket #25216 Franks 03/27/2014, UMW needs REG as a rule
                rsAT!Rule = "REG"
            Else
                rsAT!Rule = shiftinfo(0)
            End If
    
            'If glbWFC Then 'Woodbridge North Carolina
            If glbWFC Or glbMitchellPlastics Then 'Ticket #24112 Franks 07/30/2013
                If xPaidType = "Unpaid" Then
                    'rsAT!thist_start_time = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
                    'rsAT!o_in = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
                    rsAT!thist_start_time = DateAdd("n", (xTotHrs * 60), "12:00 AM")
                    rsAT!o_in = DateAdd("n", (xTotHrs * 60), "12:00 AM")
                End If '
                If xPaidType = "Paid" Then
                    ''Ticket #24439 Franks 10/01/2013 - add thist_start_time for Paid too
                    'rsAT!thist_start_time = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
                    'rsAT!thist_stop_time = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
                    'rsAT!o_out = DateAdd("n", (rsIHR!AD_HRS * 60), "12:00 AM")
                    rsAT!thist_start_time = DateAdd("n", (xTotHrs * 60), "12:00 AM")
                    rsAT!thist_stop_time = DateAdd("n", (xTotHrs * 60), "12:00 AM")
                    rsAT!o_out = DateAdd("n", (xTotHrs * 60), "12:00 AM")
                End If
                'If Len(rsIHR!AD_COMM) > 0 Then
                '    'rsAT!notes = rsIHR!AD_COMM & " - Posted From INFO:HR on " & Now
                '    rsAT!notes = rsIHR!AD_COMM & " - added by Info HR on on " & Now
                'Else
                '    rsAT!notes = "added by Info HR on " & Now
                'End If
            Else
                If Not glbPaidHours Then 'Ticket #14739
                    'rsAT!thist_start_time = DateAdd("n", rsIHR!AD_HRS * 60, "12:00 AM")
                    rsAT!thist_start_time = DateAdd("n", xTotHrs * 60, "12:00 AM")
                Else
                    rsAT!thist_start_time = shiftinfo(1)
                End If
                If glbPaidHours Then 'Ticket #14739
                    If IsDate(rsAT!thist_start_time) Then
                        'rsAT!thist_stop_time = DateAdd("n", rsIHR!AD_HRS * 60, rsAT!thist_start_time)
                        rsAT!thist_stop_time = DateAdd("n", xTotHrs * 60, rsAT!thist_start_time)
                    End If
                End If
                'rsAT!notes2 = "Posted From INFO:HR on " & Now
            End If
            If Len(rsIHR!AD_COMM) > 0 Then rsAT!notes = rsIHR!AD_COMM
            rsAT!thist_status = "C"
            rsAT.Update
            
        End If
    End If
    ' type - M ----- end
    Exit Sub
errfalse:
    'MsgBox Err.Description
    
End Sub


Public Function getRateAT(xEMP, xMultiPayCode) 'Ticket #24124 Franks 07/24/2013
Dim cnRate As New ADODB.Connection
Dim rsRate As New ADODB.Recordset
Dim ATRateConnectionString
Dim xVersion, xTable, SQLQ
On Error Resume Next

getRateAT = "0"
ATRateConnectionString = OtherDatabaseInte("Advanced Tracker", xMultiPayCode)
If ATRateConnectionString = "" Then Exit Function
If glbWFC Then 'Woodbridge North Carolina
    If InStr(ATRateConnectionString, "SQLOLEDB") <> 0 Then
        xVersion = "SQL"
        xTable = "[" & TableNamePrefix(xMultiPayCode) & "-etp.etp.rates]"
        cnRate.CursorLocation = adUseClient
        cnRate.Open ATRateConnectionString
    Else
        xVersion = "Access"
        xTable = "rates"
        cnRate.Open ATRateConnectionString
    End If
    SQLQ = "SELECT Rate FROM " & xTable & " WHERE EMP_CODE='" & xEMP & "'"
    SQLQ = SQLQ & " ORDER BY [date] DESC" 'Ticket #26960 Frank 04/16/2015
    rsRate.Open SQLQ, cnRate, adOpenStatic, adLockOptimistic
    If Not rsRate.EOF Then
        If IsNumeric(rsRate("Rate")) Then
            getRateAT = rsRate("Rate")
        End If
    End If
Else
    If InStr(ATRateConnectionString, "SQLOLEDB") <> 0 Then
        xVersion = "SQL"
        xTable = "[" & TableNamePrefix(xMultiPayCode) & "-etp.etp.employee]"
        cnRate.CursorLocation = adUseClient
        cnRate.Open ATRateConnectionString
    Else
        xVersion = "Access"
        xTable = "employee"
        cnRate.Open ATRateConnectionString
    End If
    
    SQLQ = "SELECT EMP_RATE FROM " & xTable & " WHERE EMP_CODE='" & xEMP & "'"
    rsRate.Open SQLQ, cnRate, adOpenStatic, adLockOptimistic
    If Not rsRate.EOF Then
        If IsNumeric(rsRate("EMP_RATE")) Then
            getRateAT = rsRate("EMP_RATE")
        End If
    End If
    rsRate.Close
End If
End Function

Public Function getShiftInfo(xSHIFT, xMultiPayCode)
Dim cnShift As New ADODB.Connection
Dim rsShift As New ADODB.Recordset
Dim ATShiftConnectionString
Dim xVersion, xTable, SQLQ
On Error Resume Next

getShiftInfo = "|8:00 AM"
ATShiftConnectionString = OtherDatabaseInte("Advanced Tracker", xMultiPayCode)
If ATShiftConnectionString = "" Then Exit Function


If InStr(ATShiftConnectionString, "SQLOLEDB") <> 0 Then
    xVersion = "SQL"
    xTable = "[" & TableNamePrefix(xMultiPayCode) & "-etp.etp.shift]"
    cnShift.CursorLocation = adUseClient
    cnShift.Open ATShiftConnectionString
Else
    xVersion = "Access"
    xTable = "[shift]"
    cnShift.Open ATShiftConnectionString
End If

SQLQ = "SELECT * FROM " & xTable & " WHERE SHIFT_CODE='" & xSHIFT & "'"
rsShift.Open SQLQ, cnShift, adOpenStatic, adLockOptimistic
If Not rsShift.EOF Then
    If IsDate(rsShift("SHIFT_START")) Then
        getShiftInfo = rsShift("shift_rule_num") & "|" & rsShift("shift_start")
    End If
End If
End Function

Public Function IsWFCNGSDiv(xDIV) 'Ticket #24620 Franks 12/03/2013
Dim rsNGS As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    If Len(xDIV) > 0 Then
        SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
        rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsNGS.EOF Then
            retVal = True
        End If
        rsNGS.Close
    End If
    IsWFCNGSDiv = retVal
End Function

Public Function getNGSFieldFromMatrix(xDIV, xUnion, xEmpStatus, xReturnField) 'Ticket #23247 Franks 07/23/2013
Dim rsNGS As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As String
    retVal = ""
    If glbNGS_OnFlag Then
        If Len(xDIV) > 0 And Len(xUnion) > 0 And Len(xEmpStatus) > 0 Then
            'with Status code
            SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
            SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
            SQLQ = SQLQ & "AND NG_PLAN_CODE = '" & xEmpStatus & "' "
            rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
            
            If rsNGS.EOF Then 'Ticket #23564 Franks 04/17/2013
            'check "-" status, such as "-ACT2", convert "-ACT2" to "ACT2" then compare ED_EMP with not equal to
                SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
                SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
                SQLQ = SQLQ & "AND LEFT(NG_PLAN_CODE,1) = '-' " 'for "-" code only
                SQLQ = SQLQ & "AND NOT ((CASE LEFT(NG_PLAN_CODE,1) WHEN '-' THEN REPLACE(NG_PLAN_CODE,'-', '') ELSE '' END) = '" & xEmpStatus & "') " 'convert "-ACT2" to "ACT2"; no "-" then ""
                If rsNGS.State <> 0 Then rsNGS.Close
                rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
                'if not found then without Status code
                If rsNGS.EOF Then
                    SQLQ = "SELECT * FROM WFC_NGS_SUBGROUP WHERE NG_DIV = '" & xDIV & "' "
                    SQLQ = SQLQ & "AND NG_ORG = '" & xUnion & "' "
                    SQLQ = SQLQ & "AND ((NG_PLAN_CODE IS NULL) OR NOT( NG_PLAN_CODE ='" & xEmpStatus & "')) "
                    If rsNGS.State <> 0 Then rsNGS.Close
                    rsNGS.Open SQLQ, gdbAdoIhr001, adOpenStatic
                End If
            End If
            If Not rsNGS.EOF Then
                If xReturnField = "NG_SUB_GROUP" Then retVal = rsNGS("NG_SUB_GROUP")
                If xReturnField = "NG_PAY_GROUP" Then retVal = rsNGS("NG_PAY_GROUP")
                If xReturnField = "NG_BENEFIT_GROUP" Then retVal = rsNGS("NG_BENEFIT_GROUP")
            End If
        End If
    End If
    getNGSFieldFromMatrix = retVal
End Function

Public Sub HRSoftAction(rsHRSoft As ADODB.Recordset)
    If rsHRSoft.EOF Then Exit Sub
    glbCandidate = rsHRSoft("SF_CANDIDATE")
    glbCand_SF_ID = rsHRSoft("SF_ID")
    
    If rsHRSoft("SF_HIRETYPE") = "NEW" Then
        If gSec_Add_NewHire Then
            Call WFCProNewHire(rsHRSoft)
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
        Exit Sub
    End If
    If rsHRSoft("SF_HIRETYPE") = "REH" Then
        'Ticket #24421 Franks 10/082013
        'Rehire:"   Security: User must have the "New Hire" security item checked in order to do this function.
        If gSec_Add_NewHire Then
            Call WFCProReHire(rsHRSoft)
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
        Exit Sub
    End If
    If rsHRSoft("SF_HIRETYPE") = "TRNI" Then
        If gSec_Inq_Terminations Then
            Call WFCProTransfer(rsHRSoft)
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
        Exit Sub
    End If
    If rsHRSoft("SF_HIRETYPE") = "PROM" Or rsHRSoft("SF_HIRETYPE") = "LATM" Then
        If gSec_Add_NewHire Then ' gSec_Inq_Terminations Then
            Call WFCProPROM_LATM(rsHRSoft)
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
        Exit Sub
    End If
    
    If UCase(rsHRSoft("SF_HIRETYPE")) = "UNHIRES" Then
        If gSec_Upd_Terminations Then
            Call WFCProUNHIRES(rsHRSoft)
        Else
            MsgBox "You Do Not Have Authority For This Transaction"
        End If
        Exit Sub
    End If
End Sub

Public Sub WFCProNewHire(rsHRSoft As ADODB.Recordset)
    Call UnloadFrms
    glbHRSoftType = "NewHire"
    If IsNull(rsHRSoft("SF_UPT_DEMO")) Or rsHRSoft("SF_UPT_DEMO") = 0 Then
        frmNewEmployee.Show 1
        If glbHRSoftAction = "NewEmp" Then
            frmEEBASIC.ChangeAction = NewRecord
            Call frmEEBASIC.cmdNew_Click
            Exit Sub
        End If
        Exit Sub
    End If
    'find the matching ihr emp # and name
    
    If Not FoundLocEEFind Then
        Exit Sub
    End If
    
    If IsNull(rsHRSoft("SF_UPT_STATUS")) Or rsHRSoft("SF_UPT_STATUS") = 0 Then
        Call get_WFCNewHireForms
        Do While NewHireForms.count > 0
            NewHireForms.Remove 1
            If NewHireForms.Item(1) = "frmEESTATS" Then
                Call LoadForm(NewHireForms(1))
                Exit Sub
            End If
        Loop
        Exit Sub
    End If
    If IsNull(rsHRSoft("SF_UPT_POSITION")) Or rsHRSoft("SF_UPT_POSITION") = 0 Then
        Call get_WFCNewHireForms
        Do While NewHireForms.count > 0
            NewHireForms.Remove 1
            If NewHireForms.Item(1) = "frmEPOSITION" Then
                Call LoadForm(NewHireForms(1))
                Exit Sub
            End If
        Loop
        Exit Sub
    End If
    If IsNull(rsHRSoft("SF_UPT_SALARY")) Or rsHRSoft("SF_UPT_SALARY") = 0 Then
        Call get_WFCNewHireForms
        Do While NewHireForms.count > 0
            NewHireForms.Remove 1
            If NewHireForms.Item(1) = "frmESALARY" Then
                Call LoadForm(NewHireForms(1))
                Exit Sub
            End If
        Loop
        Exit Sub
    End If
End Sub

Public Sub WFCProPROM_LATM(rsHRSoft As ADODB.Recordset)
Dim xEmpNo
Dim xRetNo
Dim xMsg, a%
    If Not IsNull(rsHRSoft("SF_EMPNBR")) Then
        If rsHRSoft("SF_HIRETYPE") = "PROM" Then glbHRSoftType = "PROM"
        If rsHRSoft("SF_HIRETYPE") = "LATM" Then glbHRSoftType = "LATM"
        
        xEmpNo = rsHRSoft("SF_EMPNBR")
        xRetNo = isValidWFCEmpNo(xEmpNo, rsHRSoft("SF_DIV"), rsHRSoft("SF_ORG"))
        If xRetNo = 0 Then
            MsgBox "Cannot find employee #" & xEmpNo & " in info:HR. Please select an employee from Active Employee list"
            glbLEE_ID = 0
            frmEEFIND.Show 1
            xEmpNo = glbLEE_ID
            If glbLEE_ID > 0 Then
                xRetNo = isValidWFCEmpNo(xEmpNo, rsHRSoft("SF_DIV"), rsHRSoft("SF_ORG"))
            Else
                Exit Sub
            End If
        End If
        'If isValidWFCEmpNo(xEmpNo) Then
        If xRetNo > 0 Then 'found employee in info:HR
        
            If Not FoundLocEEFind(xEmpNo) Then
                Exit Sub
            End If
            
            If Not WFCNamesMatched(glbCandidate, glbLEE_SName, glbLEE_FName) Then
                xMsg = "Employee Names from HRsoft does not match info:HR. " & Chr(10)
                xMsg = xMsg & Chr(10) & "   Click Yes to Use the Employee Names from info:HR  " & Chr(10)
                xMsg = xMsg & Chr(10) & "   Click No to cancel and investigate "
                a% = MsgBox(xMsg, 36, "Confirm")
                If a% <> 6 Then 'Exit Sub
                    Exit Sub
                End If
            End If
                
            '"   If the incoming candidate's plant and union code is the same, it's a promotion. Otherwise, it's a transfer.
            If xRetNo = 3 Then 'PROM & LATM
                'MsgBox "Cannot transfer this employee because the incoming candidate's division and union code is the same."
                Call UnloadFrms
                Unload frmSFFind
                'glbHRSoftType = "PROM" ' "PromLatm"

                
                Screen.MousePointer = HOURGLASS
                Load frmEWFCProm
                frmEWFCProm.ZOrder 0
                Screen.MousePointer = DEFAULT
                
                Exit Sub
            End If
            If xRetNo = 2 Then
                xMsg = "The incoming candidate's division and union code are not the same as current data. "
                xMsg = xMsg & "You have to do employee Transfer instead of Promotion/Lateral Move"
                MsgBox xMsg
                Call UnloadFrms
                Unload frmSFFind
                glbHRSoftType = "Transfer"

                glbTermTran = False

    
                Screen.MousePointer = HOURGLASS
                Load frmETRANIN
                frmETRANIN.ZOrder 0
                Screen.MousePointer = DEFAULT
                
                Exit Sub
            End If
        Else
            '' = 0
            'MsgBox "Cannot find employee #" & xEmpNo & " in info:HR"
            'Exit Sub
        End If
    End If

End Sub

Public Sub WFCProUNHIRES(rsHRSoft As ADODB.Recordset)
Dim rsAudit As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim xEmpNo As Long
Dim xRetNo
Dim xMsg As String, a%
Dim xSFID
Dim xFoundInHRSoft As Boolean
    '"   Use the Candidate ID field to lookup the record in the HRSOFT file
    '"   If not found, delete the record from the HRSOFT file
    xSFID = rsHRSoft("SF_ID")
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE NOT SF_ID = " & rsHRSoft("SF_ID") & " "
    SQLQ = SQLQ & "AND SF_CANDIDATE = " & rsHRSoft("SF_CANDIDATE") & " "
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsTemp.EOF Then
        xFoundInHRSoft = False
    Else
        xFoundInHRSoft = True
    End If
    If Not xFoundInHRSoft Then
        'xMsg = "No any none UNHIRES record in HRSOFT file. "
        'xMsg = xMsg & Chr(10) & "Do you want to delete the UNHIRES record? "
        xMsg = "Do you want to delete this UNHIRE record?"
        a% = MsgBox(xMsg, 36, "Confirm Delete")
        If a% = 6 Then 'Yes
            rsTemp.Close
            SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE SF_ID = " & rsHRSoft("SF_ID") & " "
            gdbAdoIhr001.Execute SQLQ
            glbCandidate = 0
        End If
        Exit Sub
    End If
    If rsTemp.State <> 0 Then rsTemp.Close
    
    If xFoundInHRSoft Then
        ' If found and the Process Flag equals "Y", lookup the Employee Master via the Candidate ID
        SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE NOT SF_ID = " & rsHRSoft("SF_ID") & " "
        SQLQ = SQLQ & "AND SF_CANDIDATE = " & rsHRSoft("SF_CANDIDATE") & " "
        'SQLQ = SQLQ & "AND (SF_HIRETYPE = 'NEW' OR SF_HIRETYPE = 'REH' OR) AND  "
        SQLQ = SQLQ & " AND NOT (SF_UPT_PROCESSED = 0) " 'Process Flag equals "Y"
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            'lookup the Employee Master via the Candidate ID.
            If Not FoundLocEEFind Then
                ''"   If not found, delete the record from the HRSOFT file
                'xMsg = "Found processed record in HRSOFT file for Candidate #" & rsHRSoft("SF_CANDIDATE") & ", but there is no matching record in info:HR table. "
                'xMsg = xMsg & Chr(10) & "The program is going to delete all records for Candidate #" & rsHRSoft("SF_CANDIDATE") & " in HRSOFT file "
                'xMsg = xMsg & Chr(10) & "Do you want to do this? "
                
                'Ticket #24936 Franks 02/04/2014
                xMsg = "This employee's New Hire/Rehire function has been initiated. "
                xMsg = xMsg & Chr(10) & "Do you want to undo the info:HR transactions and delete the UNHIRE record?"
                a% = MsgBox(xMsg, 36, "Confirm Delete")
                If a% = 6 Then 'Yes
                    rsTemp.Close
                    SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE SF_ID = " & rsHRSoft("SF_ID") & " "
                    gdbAdoIhr001.Execute SQLQ
                    glbCandidate = 0
                End If
                Exit Sub
            Else
                'found match record in HREMP
                xEmpNo = glbLEE_ID
                '"   If found, lookup the audit master's au_upload flag.
                SQLQ = "SELECT * FROM HRAUDIT WHERE AU_EMPNBR = " & xEmpNo & " AND AU_UPLOAD = 'Y' "
                If rsAudit.State <> 0 Then rsAudit.Close
                rsAudit.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsAudit.EOF Then
                    'found record of AU_UPLOAD = 'Y',
                    'o   If yes, update the HRAUDIT with a termination record. Use today's date as the date of termination and Termination Reason as EEIN.
                    '"   Perform a termination process for this employee.
                    xMsg = ""
                    xMsg = xMsg & Chr(10) & "The program is going to terminate this employee "
                    xMsg = xMsg & Chr(10) & "The employee # is " & xEmpNo & " "
                    xMsg = xMsg & Chr(10) & Chr(10) & "Do you want to do this? "
                    a% = MsgBox(xMsg, 36, "Termination")
                    If a% = 6 Then 'Yes
                        rsAudit.Close
                        'Call WFCEmployeeDele(xEmpNo)
                        Screen.MousePointer = HOURGLASS
                        Call WFCHRSoftTermination(xEmpNo, Date, glbEmpDiv, glbUNION, True)
                        'SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE NOT SF_ID = " & rsHRSoft("SF_ID") & " "
                        'gdbAdoIhr001.Execute SQLQ
                        
                        Call WFCHRSoftProcUpt("UNHIRES")
                        glbCandidate = 0
                        MDIMain.panHelp(0).FloodType = 0
                        Screen.MousePointer = DEFAULT
                        MsgBox "   Finished.   "
                    End If
                Else
                    'No record of AU_UPLOAD = 'Y',
                    'delete all audit records for this employee and delete all records in all tables for this employee number (includes NGS, MLF and Pension).
                    xMsg = "No record with AU_UPLOAD = 'Y' in info:HR Audit table for this employee "
                    xMsg = xMsg & Chr(10) & "The program is going to delete all audit records for this employee and delete all records in all tables for this employee number (includes NGS, MLF and Pension). "
                    xMsg = xMsg & Chr(10) & "The employee # is " & xEmpNo & " "
                    xMsg = xMsg & Chr(10) & Chr(10) & "Do you want to do this? "
                    a% = MsgBox(xMsg, 36, "Warning - No Recovery Delete!")
                    If a% = 6 Then 'Yes
                        rsAudit.Close
                        Call WFCEmployeeDele(xEmpNo)
                        'SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE NOT SF_ID = " & rsHRSoft("SF_ID") & " "
                        SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE SF_ID = " & rsHRSoft("SF_ID") & " "
                        gdbAdoIhr001.Execute SQLQ
                        glbCandidate = 0
                    End If
                End If
                Exit Sub
            End If
    
        End If
        
        'If found and the Hire Type equals "NEW" or "REH" and the Process Flag equals "N", delete all candidate records in the HRSOFT file including the unhire record.
        SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE NOT SF_ID = " & rsHRSoft("SF_ID") & " "
        SQLQ = SQLQ & "AND SF_CANDIDATE = " & rsHRSoft("SF_CANDIDATE") & " "
        SQLQ = SQLQ & "AND (SF_HIRETYPE = 'NEW' OR SF_HIRETYPE = 'REH') "
        SQLQ = SQLQ & " AND (SF_UPT_PROCESSED = 0) "
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            xMsg = "Found unprocessed 'NEW' or 'REH' record in HRSOFT file. "
            xMsg = xMsg & Chr(10) & "The program is going to delete all records for Candidate #" & rsHRSoft("SF_CANDIDATE") & " "
            xMsg = xMsg & Chr(10) & "Do you want to do this? "
            a% = MsgBox(xMsg, 36, "Confirm Delete")
            If a% = 6 Then 'Yes
                rsTemp.Close
                SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE SF_ID = " & rsHRSoft("SF_ID") & " "
                gdbAdoIhr001.Execute SQLQ
                glbCandidate = 0
            End If
            Exit Sub
        Else
            SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE NOT SF_ID = " & rsHRSoft("SF_ID") & " "
            SQLQ = SQLQ & "AND SF_CANDIDATE = " & rsHRSoft("SF_CANDIDATE") & " "
            'SQLQ = SQLQ & "AND NOT (SF_HIRETYPE = 'NEW' OR SF_HIRETYPE = 'REH') "
            SQLQ = SQLQ & "AND (SF_UPT_PROCESSED = 0) "
            If rsTemp.State <> 0 Then rsTemp.Close
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                xMsg = "Found unprocessed record in HRSOFT file. "
                xMsg = xMsg & Chr(10) & "The program is going to delete all records for Candidate #" & rsHRSoft("SF_CANDIDATE") & " "
                xMsg = xMsg & Chr(10) & "Do you want to do this? "
                a% = MsgBox(xMsg, 36, "Confirm Delete")
                If a% = 6 Then 'Yes
                    rsTemp.Close
                    SQLQ = "DELETE FROM HRSF_XML_IMPORT WHERE SF_ID = " & rsHRSoft("SF_ID") & " "
                    gdbAdoIhr001.Execute SQLQ
                    glbCandidate = 0
                End If

            End If
            Exit Sub
        End If
    End If
    'check if
    
End Sub

Public Sub WFCProTransfer(rsHRSoft As ADODB.Recordset)
Dim xEmpNo
Dim xRetNo
Dim xMsg
    If Not IsNull(rsHRSoft("SF_EMPNBR")) Then
        xEmpNo = rsHRSoft("SF_EMPNBR")
        xRetNo = isValidWFCEmpNo(xEmpNo, rsHRSoft("SF_DIV"), rsHRSoft("SF_ORG"))
        
        If xRetNo = 0 Then
            MsgBox "Cannot find employee #" & xEmpNo & " in info:HR. Please select an employee from Active Employee list"
            glbLEE_ID = 0
            frmEEFIND.Show 1
            xEmpNo = glbLEE_ID
            If glbLEE_ID > 0 Then
                xRetNo = isValidWFCEmpNo(xEmpNo, rsHRSoft("SF_DIV"), rsHRSoft("SF_ORG"))
            Else
                Exit Sub
            End If
        End If
        
        'If isValidWFCEmpNo(xEmpNo) Then
        If xRetNo > 0 Then
        
            If Not FoundLocEEFind(xEmpNo) Then
                Exit Sub
            End If
            
            If Not WFCNamesMatched(glbCandidate, glbLEE_SName, glbLEE_FName) Then
                xMsg = "Employee Names from HRsoft does not match info:HR. " & Chr(10)
                xMsg = xMsg & Chr(10) & "   Click Yes to Use the Employee Names from info:HR  " & Chr(10)
                xMsg = xMsg & Chr(10) & "   Click No to cancel and investigate "
                a% = MsgBox(xMsg, 36, "Confirm")
                If a% <> 6 Then 'Exit Sub
                    Exit Sub
                End If
            End If
            
            '"   If the incoming candidate's plant and union code is the same, it's a promotion. Otherwise, it's a transfer.
            If xRetNo = 1 Then
                MsgBox "Cannot transfer this employee because the incoming candidate's division and union code is the same."
                Exit Sub
                
            End If
            If xRetNo = 2 Then
                Call UnloadFrms
                Unload frmSFFind
                glbHRSoftType = "Transfer"

                glbTermTran = False
    
                Screen.MousePointer = HOURGLASS
                Load frmETRANIN
                frmETRANIN.ZOrder 0
                Screen.MousePointer = DEFAULT
                
                Exit Sub
            End If
            If xRetNo = 3 Then 'PROM & LATM
                'MsgBox "Cannot transfer this employee because the incoming candidate's division and union code is the same."
                xMsg = "The incoming candidate's division and union code are the same as current data. "
                xMsg = xMsg & "You have to do employee Promotion instead of Transfer"
                MsgBox xMsg
                Call UnloadFrms
                Unload frmSFFind
                glbHRSoftType = "PROM" '
                
                Screen.MousePointer = HOURGLASS
                Load frmEWFCProm
                frmEWFCProm.ZOrder 0
                Screen.MousePointer = DEFAULT
                
                Exit Sub
            End If
        Else
            'MsgBox "Cannot find employee #" & xEmpNo & " in info:HR"
            Exit Sub
        End If
    End If
End Sub

Public Sub WFCProReHire(rsHRSoft As ADODB.Recordset)
    Call UnloadFrms
    Unload frmSFFind
    glbHRSoftType = "ReHire"
    'If IsNull(rsHRSoft("SF_UPT_DEMO")) Or rsHRSoft("SF_UPT_DEMO") = 0 Then
    If rsHRSoft("SF_UPT_REHIRE") = 0 Then
        'frmEEBASIC.ChangeAction = NewRecord
        'Call frmEEBASIC.cmdNew_Click
        frmNewEmployee.Show 1
        If glbHRSoftAction = "NewEmp" Then
            frmEEBASIC.ChangeAction = NewRecord
            Call frmEEBASIC.cmdNew_Click
            Exit Sub
        End If
        If glbHRSoftAction = "ReHireEmp" Then
            Screen.MousePointer = HOURGLASS
            Load frmEREHIRE
            Call frmEREHIRE.SET_UP_MODE
            frmEREHIRE.ZOrder 0
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
        Exit Sub
    End If
    'find the matching ihr emp # and name
    
    If Not FoundLocEEFind Then
        Exit Sub
    End If
    If IsNull(rsHRSoft("SF_UPT_DEMO")) Or rsHRSoft("SF_UPT_DEMO") = 0 Then
        Call get_WFCNewHireFormShort 'modify here to have the 5 screen only
        Call LoadForm("frmEEBASIC")
        'Do While NewHireForms.count > 0
        '    'NewHireForms.Remove 1
        '    'If NewHireForms.Item(1) = "frmEESTATS" Then
        '    '    Call LoadForm(NewHireForms(1))
        '    '    Exit Sub
        '    'End If
        'Loop
        Exit Sub
    End If
    If IsNull(rsHRSoft("SF_UPT_STATUS")) Or rsHRSoft("SF_UPT_STATUS") = 0 Then
        Call get_WFCNewHireFormShort 'modify here to have the 5 screen only
        Do While NewHireForms.count > 0
            NewHireForms.Remove 1
            If NewHireForms.Item(1) = "frmEESTATS" Then
                Call LoadForm(NewHireForms(1))
                Exit Sub
            End If
        Loop
        Exit Sub
    End If
    If IsNull(rsHRSoft("SF_UPT_POSITION")) Or rsHRSoft("SF_UPT_POSITION") = 0 Then
        Call get_WFCNewHireFormShort
        Do While NewHireForms.count > 0
            NewHireForms.Remove 1
            If NewHireForms.Item(1) = "frmEPOSITION" Then
                Call LoadForm(NewHireForms(1))
                Exit Sub
            End If
        Loop
        Exit Sub
    End If
    If IsNull(rsHRSoft("SF_UPT_SALARY")) Or rsHRSoft("SF_UPT_SALARY") = 0 Then
        Call get_WFCNewHireFormShort
        Do While NewHireForms.count > 0
            NewHireForms.Remove 1
            If NewHireForms.Item(1) = "frmESALARY" Then
                Call LoadForm(NewHireForms(1))
                Exit Sub
            End If
        Loop
        Exit Sub
    End If
'    'banking
End Sub

Public Function FoundLocEEFind(Optional xEmpNo = 0)
Dim rsTmp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = True
    If xEmpNo = 0 Then
        SQLQ = "SELECT * FROM HREMP WHERE ED_CANDIDATE = " & glbCandidate
    Else
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    End If
    rsTmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTmp.EOF Then
        glbLEE_ID = rsTmp("ED_EMPNBR")
        glbEmpCountry = UCase(rsTmp("ED_COUNTRY"))
        
        If IsNull(rsTmp("ED_ORG")) Then
            glbUNION = ""
        Else
            glbUNION = rsTmp("ED_ORG")
        End If
        If Not IsNull(rsTmp("ED_FNAME")) Then
            glbLEE_FName = rsTmp("ED_FNAME")
        Else
            glbLEE_FName = "*ERROR*"
        End If
        If Not IsNull(rsTmp("ED_SURNAME")) Then
            glbLEE_SName = rsTmp("ED_SURNAME")
        Else
            glbLEE_SName = "*ERROR*"
        End If
        If glbWFC Then 'Get the glbBand
            glbBand = get_WFCband(glbLEE_ID)
            If IsNull(rsTmp("ED_SIN")) Then 'Ticket #18566
                glbSIN = ""
            Else
                glbSIN = rsTmp("ED_SIN")
            End If
            'Ticket #19266 - BEGIN
            If IsNull(rsTmp("ED_VADIM1")) Then
                glbWFCNGSSubGroup = ""
            Else
                glbWFCNGSSubGroup = rsTmp("ED_VADIM1")
            End If
            If IsNull(rsTmp("ED_VADIM2")) Then
                glbWFCPayGroup = ""
            Else
                glbWFCPayGroup = rsTmp("ED_VADIM2")
            End If
            glbEmpDiv = rsTmp("ED_DIV")
            'Ticket #19266 - END
        End If
    Else
        retVal = False
    End If

    FoundLocEEFind = retVal
End Function
Public Sub get_WFCNewHireFormShort()
Dim xFormItem(5)
xFormItem(1) = "frmEEBASIC"
xFormItem(2) = "frmEESTATS"
xFormItem(3) = "frmEPOSITION"
xFormItem(4) = "frmESALARY"
xFormItem(5) = "frmEBANK"

For X = 1 To NewHireForms.count: NewHireForms.Remove 1: Next
For X = 1 To 5
    NewHireForms.Add Trim(xFormItem(X))
Next
End Sub
Public Sub get_WFCNewHireForms()
Dim rsTN As New ADODB.Recordset, X
rsTN.Open "SELECT * FROM HRNEWHIRE WHERE NewHire<>0 ORDER BY ID", glbAdoIHRDB
For X = 1 To NewHireForms.count: NewHireForms.Remove 1: Next
Do Until rsTN.EOF
    'If Trim(strNoAccessForms) <> "" Then
    '    'Skip the screens the user does not have access to
    '    If InStr(strNoAccessForms, Trim(rsTN!MenuItem)) = 0 Then
    '        NewHireForms.Add Trim(rsTN!FormName)
    '    End If
    'Else
        NewHireForms.Add Trim(rsTN!FormName)
    'End If
    rsTN.MoveNext
Loop
rsTN.Close
End Sub

Private Function get_WFCband(empNo)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
    get_band = ""
    SQLQ = "SELECT SH_EMPNBR,SH_BAND FROM HR_SALARY_HISTORY WHERE SH_CURRENT <>0 AND SH_EMPNBR = " & empNo
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("SH_BAND")) Then
            get_band = rsTemp("SH_BAND")
        End If
    End If
    rsTemp.Close
End Function

Public Function IsWFCNotInTracker(xEmpNo) 'Ticket #27609 Franks 10/07/2015
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT ED_EMPNBR,ED_BONUSDEPT FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        If Not IsNull(rsTemp("ED_BONUSDEPT")) Then
            If rsTemp("ED_BONUSDEPT") = "000000" Then
                retVal = True
            End If
        End If
    End If
    IsWFCNotInTracker = retVal
End Function

Public Function IsWFCAdvConnected()
Dim ATConnectionString
Dim cnAT As New ADODB.Connection
Dim xTable
Dim retVal As Boolean
On Error GoTo conn_error
    retVal = True
    If Len(glbPlantCode) > 0 Then
        If Not UCase(glbPlantCode) = "ALL" Then
            ATConnectionString = OtherDatabaseInte("Advanced Tracker", glbPlantCode)
            If Len(ATConnectionString) = 0 Then
                retVal = False
            Else
                'xTable = "[" & TableNamePrefix(xSection) & "-etp.etp.employee]"
                cnAT.CursorLocation = adUseClient
                cnAT.Open ATConnectionString
                cnAT.Close
            End If
        End If
    End If
    IsWFCAdvConnected = retVal
    Exit Function
conn_error:
    retVal = False
    IsWFCAdvConnected = retVal
End Function

Public Function isValidWFCEmpNo(xEmpNo, xDIV, xORG)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retVal
    retVal = 0
    SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_ORG FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retVal = 1 'valid EmpNo
        If Not IsNull(xDIV) And Not IsNull(xORG) Then
            If Not IsNull(rsTemp("ED_DIV")) And Not IsNull(rsTemp("ED_ORG")) Then
                If Not (xDIV = rsTemp("ED_DIV") And xORG = rsTemp("ED_ORG")) Then
                    retVal = 2 'Div or Union changed
                Else
                    retVal = 3 'Div or Union not changed
                End If
            End If
        End If
    End If
    rsTemp.Close
    isValidWFCEmpNo = retVal
End Function

Public Sub WFCPosSkillsUpd(xEmpNo, xJobCode, xStartDate)
Dim rsMain As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsEmpSki As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HRJOBSKL WHERE JS_CODE = '" & xJobCode & "' "
    SQLQ = SQLQ & "AND JS_EXPFACT = 0 "
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsMain.EOF
        'open another record setup with all records for this skill, insert it to employee skill table
        SQLQ = "SELECT * FROM HRJOBSKL WHERE JS_CODE = '" & xJobCode & "' "
        SQLQ = SQLQ & "AND JS_SKILL = '" & rsMain("JS_SKILL") & "' "
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsTemp.EOF
            SQLQ = "SELECT * FROM HREMPSKL WHERE SE_EMPNBR = " & xEmpNo & " "
            SQLQ = SQLQ & "AND SE_SKILL = '" & rsTemp("JS_SKILL") & "' " '
            SQLQ = SQLQ & "AND SE_LEVEL = " & rsTemp("JS_EXPFACT") & " "
            If rsEmpSki.State <> 0 Then rsEmpSki.Close
            rsEmpSki.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If rsEmpSki.EOF Then
                rsEmpSki.AddNew
                rsEmpSki("SE_COMPNO") = "001"
                rsEmpSki("SE_EMPNBR") = xEmpNo
                rsEmpSki("SE_SKILL") = rsTemp("JS_SKILL")
                rsEmpSki("SE_LEVEL") = rsTemp("JS_EXPFACT")
                rsEmpSki("SE_DATE") = xStartDate
                rsEmpSki("SE_LDATE") = Date
                rsEmpSki("SE_LTIME") = Time$
                rsEmpSki("SE_LUSER") = glbUserID
            End If
            rsEmpSki.Update
            rsTemp.MoveNext
        Loop
        rsMain.MoveNext
    Loop
    rsMain.Close
End Sub

Public Sub WFCHRSoftTermination(xEmpNo, xTermDate, xNewDiv, xNewUnion, xIsTerm As Boolean) 'Ticket #24184 Franks 01/13/2014
Dim rsTB As New ADODB.Recordset
Dim rsEM As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim rsNGSAUDIT As New ADODB.Recordset
Dim fglbEMPNBR As Long
Dim AbortTerm As Boolean
Dim fglbFollowID
Dim fglbNew
Dim glbPicDir, glbPicBMP
Dim locCertNo As String
Dim locWFCPenEligible As Boolean
Dim locWFCPenEarnFlag As Boolean
Dim locSection As String
Dim locUnion As String
Dim locSIN As String
Dim locPayrollID As String
Dim xLocID
Dim xEmpName
Dim EID&
Dim xTermReason As String
Dim xNGSStart
Dim xCurPosition

    If xIsTerm Then
        xTermReason = "EEIN"
    Else
        xTermReason = "TOUT"
    End If
    
    xNGSStart = ""
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            xNGSStart = rsEmpOther("ER_OTHERDATE1")
        End If
    End If
    rsEmpOther.Close
    
    'xTermDate = DateAdd("D", -1, CVDate(xTranInDate))
    'chkTerms ----------------------- begin
    locCertNo = ""
    locDiv = ""
    locEmpStatus = ""
    locSection = ""
    locUnion = ""
    locSIN = ""
    locPayrollID = ""
    locBenGroup = "" 'Ticket #24176 Franks 08/07/2013
    xEmpName = ""
    xCurPosition = getEmpPostion(xEmpNo) 'Ticket #25911 Franks 12/18/2014
    
    'If IsDate(xTERMDATE) Then
    
        rsEM.Open "SELECT * FROM HREMP WHERE ED_EMPNBR=" & xEmpNo, gdbAdoIhr001, adOpenKeyset
        If rsEM.EOF Then Exit Sub
        
        If Not IsNull(rsEM("ED_USER_TEXT1")) Then
            locCertNo = rsEM("ED_USER_TEXT1")
        End If
        If Not IsNull(rsEM("ED_DIV")) Then
            locDiv = rsEM("ED_DIV")
        End If
        If Not IsNull(rsEM("ED_EMP")) Then
            locEmpStatus = rsEM("ED_EMP")
        End If
        If Not IsNull(rsEM("ED_SECTION")) Then
            locSection = rsEM("ED_SECTION")
        End If
        If Not IsNull(rsEM("ED_ORG")) Then
            locUnion = rsEM("ED_ORG")
        End If
        If Not IsNull(rsEM("ED_SIN")) Then
            locSIN = rsEM("ED_SIN")
        End If
        If Not IsNull(rsEM("ED_PAYROLL_ID")) Then
            locPayrollID = rsEM("ED_PAYROLL_ID")
        End If
        If Not IsNull(rsEM("ED_BENEFIT_GROUP")) Then locBenGroup = rsEM("ED_BENEFIT_GROUP") 'Ticket #24176 Franks 08/07/2013
        xEmpName = rsEM("ED_SURNAME") & ", " & rsEM("ED_FNAME")
        rsEM.Close
    'End If
    'chkTerms ----------------------- end
    
    glbChgBenTermDate = ""
    ''''Logic:
    ''''1. Transfer out, don't do any thing for Manulife
    ''''2. Termination, always terminate Benefits and Dependents
    '''If Not glbTermTran Then 'Transfer Out
    ''    ''glbChgBenTermDate = dlpBenCeaseDate
    ''    'Ticket #24451 Franks 10/15/2013
    ''    '"   For a MLF employee, when transferring out from Canada to another location, the benefits would have to end.
    ''    'The benefit end date should equal the Transfer Out Date
    ''    If Len(locCertNo) > 0 Then
    ''        If IsOutFromBenGrpMatrix(xNewDiv) Then
    ''            glbChgBenTermDate = xTermDate
    ''        End If
    ''    End If
    '''End If
        
    If Not UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then 'do not use it
        If gsEMAIL_ONTERM Then
            If WFCNonUnion(xEmpNo) Then
                ' Make sure we have needed info to send email
                If GetEmpData(xEmpNo, "ED_EMAIL") = "" Then ' And Not MDIMain.mnu_File_EmailSetup.Visible Then
                    Screen.MousePointer = vbDefault
                    MsgBox GetEmpData(xEmpNo, "ED_FNAME") & ", please fill in your email address on the Status/Dates screen, before attempting to terminate an employee.", vbExclamation + vbOKOnly, "Missing Email Address"
                    Exit Sub
                Else
                    If Not IsEmailSetup(xEmpNo) Then 'MDIMain.mnu_File_EmailSetup.Visible And Not IsEmailSetup(glbEmpNbr) Then  'lost condition afther removing menu items , should check
                        Screen.MousePointer = vbDefault
                        MsgBox "You have not been set up for email sending.  Please use the Setup->Security->Email Setup menu option to set up your account for email sending before attempting to terminate salaried employees.  Termination aborted.", vbCritical + vbOKOnly, "No Email Setup Found"
                        Exit Sub
                    End If
                End If
                ' Send the email
                'cmdEmail_Click
                Call WFCTermTransferEmailSending(xEmpNo, xEmpName, xTermDate, True)
                
                ' AC - dkostka - 05/03/01 - Added error checking, refuse to terminate if email didn't go through
                If AbortTerm = True Then
                    'Screen.MousePointer = vbDefault
                    'MDIMain.panHelp(0).FloodType = 1
                    'MDIMain.panHelp(0).Caption = "Termination Aborted"
                    'MsgBox "Error sending email.  Termination aborted.", vbCritical + vbOKOnly, "Error"
                    'Exit Sub
                    'Ticket #24422 Franks 10/01/2013 - "Can't stop the termination or transfer out if the email sending doesn't work.
                    MsgBox "Error sending termination email."
                End If
            End If
        End If
    End If
            
    If WFCPensionEligible(xEmpNo) And xIsTerm Then 'term
        Call WFCPensionAlerts(xEmpNo, xTermDate, "Termination - " & xTermReason, , , , "ALL")
        Call WFCPensionAlerts(xEmpNo, xTermDate, "Termination - " & xTermReason)
    End If
    
    'NGS and Benefit
    If Not xIsTerm Then 'transfer out
        Call WFC_NGS_Trans_TermTransferOut("Transfer Out", xTermDate, xNewDiv, xNewUnion, "")
        
        ''Ticket #24767 Franks 12/11/2013
        'If frmWFCBenList.Visible Then 'US NGS employees
        '    Call WFC_NGSBenEndDateUpt(glbLEE_ID)
        'End If
        'based on the new div
        If IsDate(xNGSStart) Then 'NGS employee with start date
            If Not IsWFCNGSDiv(xNewDiv) Then
                Call WFCUpdateBenefitGroup(xEmpNo)
                Call WFCUptData2fromDOT(xEmpNo, xTermDate)
                Call WFC_NGSBenEndDateUpPub(xEmpNo, "")
            End If
        End If
                    
    Else 'termination
        'If dlpDOther2.Visible Then
        If IsDate(xNGSStart) Then 'NGS employee with start date
            '"   Check the NGS Audit. If found, check the upload flag. If the upload flag is N, delete the record. If the upload flag is Y, create a new record with the NG_FORM equal to "Termination". The NGS End Date would be today's date
            SQLQ = "SELECT * FROM WFC_NGS_AUDIT WHERE NG_EMPNBR = " & xEmpNo & " AND NG_UPLOAD = 'Y' "
            If rsNGSAUDIT.State <> 0 Then rsNGSAUDIT.Close
            rsNGSAUDIT.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If rsNGSAUDIT.EOF Then
                SQLQ = "DELETE FROM WFC_NGS_AUDIT WHERE NG_EMPNBR = " & xEmpNo & " "
                gdbAdoIhr001.Execute SQLQ
            Else
                'Call WFC_NGS_Trans_TermTransferOut("Termination", xTermDate, xNewDiv, xNewUnion, xNGSStart)
                Call WFC_NGS_Trans_TermTransferOut("Termination", Date, xNewDiv, xNewUnion, xNGSStart)
            End If
        End If
        
        'Ticket #23948 Frank 06/24/2013
        Call WFC_UptPenDate4WithDOT(xEmpNo, xTermDate)
        
        ''Ticket #23247 Franks 07/22/2013
        'If frmWFCBenList.Visible Then 'US NGS employees
        '    Call WFC_NGSBenEndDateUpt(glbLEE_ID)
        'End If
        If IsDate(xNGSStart) Then 'NGS employee with start date
            Call WFCUpdateBenefitGroup(xEmpNo)
            Call WFCUptData2fromDOT(xEmpNo, xTermDate)
            Call WFC_NGSBenEndDateUpPub(xEmpNo, xTermDate)
        End If
    End If
    
    rsTB.Open "Term_HRSEQ", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If rsTB.EOF And rsTB.BOF Then
        glbTERM_Seq = 1
        rsTB.AddNew
    Else
        rsTB.MoveFirst 'Jaddy 10/28/99
        glbTERM_Seq = rsTB("TERM_SEQ_NEXT")
    End If
    rsTB("TERM_SEQ_NEXT") = glbTERM_Seq + 1
    rsTB.Update
    rsTB.Close
    
    'Ticket #25669 Franks 06/24/2014 - comment out the following code
    '''Ticket #20270 Franks 05/05/2011
    '''If glbTermTran Then
    ''If xIsTerm Then
    ''    'Call EEO_Process
    ''    If glbEmpCountry = "U.S.A." Then
    ''        Call uptEEO_Fields(glbLEE_ID, "Delete")
    ''    End If
    ''Else
    ''    'Ticket #24422 Franks 10/02/2013
    ''    'Transfer Out should not delete the EEO records.-
    ''End If
    
    If xIsTerm Then
        Call WFCAUDITTermTransferOut(xTermDate, xNewDiv, xNewUnion, True)
    Else
        Call WFCAUDITTermTransferOut(xTermDate, xNewDiv, xNewUnion, False)
    End If
    
    Call AUDIT_MANULIFE_TRANS_TermTransferOut(locCertNo)

    
    'Ticket #16395 - Pension System - begin
    If xIsTerm Then 'termination
        'Termination -
        toSOURCE = "IHR Termination" 'Ticket #19954
        'xPAData = "PA"
        'If locWFCPenEarnFlag Then
        '    If IsNumeric(medAmount.Text) Then
        '        xPAData = "PA|" & Trim(Str(medAmount.Text))
        '    End If
        'End If
        'Call WFCPensionMasUpt(glbLEE_ID, "Termination", xTERMDATE, xTermReason, Year(CVDate(xTERMDATE)), xPAData)
        Call WFCPensionMasUpt(glbLEE_ID, "Termination", xTermDate, xTermReason, Year(CVDate(xTermDate)))

        'If clpCode(1).Text = "DECD" Then
        '    'One employee can have one DBS plus other DB pensions, such as DBKIPL
        '    'Employee Dan Dubblestyne had DBS and DBKIPL pensions, create other pensions for this year with status "D"
        '    xPenType = getDBType(locSection, locUnion, "PenType")
        '    Call WFCOtherPenUpt(glbLEE_ID, glbSIN, Year(dlpTermDate.Text), "", xPenType, "D", dlpTermDate.Text, dlpTermDate.Text, "DB")
        'End If
        
        'Ticket #22009 Franks 05/09/2012
        'delete other Alerts which were created in Termination
        Call WFCPensionAlerts(glbLEE_ID, xTermDate, "Termination - " & xTermReason, , , , "Y")
            
    Else
        'Transfer Out - Trans Date, New Div
        toSOURCE = "IHR Transfer Out" 'Ticket #19954
        Call WFCPensionMasUpt(xEmpNo, "Transfer Out", xTermDate, xNewDiv, Year(CVDate(xTermDate)))
        'Ticket #19678 Franks 01/24/2011
        'On Transfer out:   (Plant Code equals "TILB" and Union Code is "C127") or (Plant Code equals "WHBY" and Union Code is "C222") the Hire Code equals "N".
        Call WFCHireCode(xEmpNo)
    
        'Ticket #21677 Franks 03/14/2012
        If xNewUnion = "NONE" Or xNewUnion = "EXEC" Then
            Call locWFCUpdPAMaster_TermTransferOut(xEmpNo, xTermDate, locSIN, locDiv, locPayrollID)
        End If
    End If
    'Ticket #16395 - Pension System - end
        
    'If Not modTermMove() Then Exit Sub
    Call modTermMoveWFCTransferOut(xTermDate, xTermReason, True)

    EID& = xEmpNo

    Call Term_Superv_General(EID&)

    Call Term_Reviewer_General(EID&)
    
    MDIMain.panHelp(0).FloodPercent = 100
    
    Call UpdEHScorrective(EID&, xEmpName)
    
    If xIsTerm Then 'termination
        Call NukeEE2_General(EID&)
    Else
        'Ticket #25927 Franks 08/25/2014
        'HRSoft transfer out, this employee still keep in HREMP
    End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~'added by RAUBREY 5/23/97 ~~~~~~~~~~~~~~~~~~~~~~
    rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsT_PARCO("PC_NUMBER_EMPLOYEES") - 1 'UPDATE FIELD WITH ACTUAL COUNT
    rsT_PARCO.Update
    rsT_PARCO.Close
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    'If Not glbTermTran Then 'Transfer Out
    If Not xIsTerm Then
        SQLQ = "UPDATE Term_HREMP SET ED_OMERS=" & Date_SQL(xTermDate)
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
        gdbAdoIhr001X.Execute SQLQ
        Call UptLUserLDateLTime(glbTERM_Seq) 'Ticket #24355 Franks 09/17/2013
    End If
    'End If

    If glbAdv Then 'Ticket #15074
        Call Employee_Master_Integration(xEmpNo, "T" & Trim(Str(glbTERM_Seq)))
    End If
    
    If Len(xCurPosition) > 0 Then
        Call mod_Upd_Pos_Budget_WFC(xCurPosition, "") 'Ticket #25911 Franks 12/18/2014
    End If
    
End Sub

Public Sub WFCHRSoftTransferOut(xNewDiv, xNewUnion, xTranInDate) 'Ticket #24184 Franks 10/28/2013
Dim rsTB As New ADODB.Recordset
Dim rsEM As New ADODB.Recordset
Dim rsT_PARCO As New ADODB.Recordset
Dim fglbEMPNBR As Long
Dim AbortTerm As Boolean
Dim fglbFollowID
Dim fglbNew
Dim glbPicDir, glbPicBMP
Dim locCertNo As String
Dim locWFCPenEligible As Boolean
Dim locWFCPenEarnFlag As Boolean
Dim locSection As String
Dim locUnion As String
Dim locSIN As String
Dim locPayrollID As String
Dim xLocID
Dim xTermDate
Dim xEmpName
Dim EID&
    'use WFCHRSoftTermination instead of WFCHRSoftTransferOut
    Exit Sub
    
    xTermDate = DateAdd("D", -1, CVDate(xTranInDate))
    'chkTerms ----------------------- begin
    locCertNo = ""
    locDiv = ""
    locEmpStatus = ""
    locSection = ""
    locUnion = ""
    locSIN = ""
    locPayrollID = ""
    locBenGroup = "" 'Ticket #24176 Franks 08/07/2013
    xEmpName = ""
    If IsDate(xTermDate) Then
    
        rsEM.Open "SELECT * FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
        If rsEM.EOF Then Exit Sub
        
        If Not IsNull(rsEM("ED_USER_TEXT1")) Then
            locCertNo = rsEM("ED_USER_TEXT1")
        End If
        If Not IsNull(rsEM("ED_DIV")) Then
            locDiv = rsEM("ED_DIV")
        End If
        If Not IsNull(rsEM("ED_EMP")) Then
            locEmpStatus = rsEM("ED_EMP")
        End If
        If Not IsNull(rsEM("ED_SECTION")) Then
            locSection = rsEM("ED_SECTION")
        End If
        If Not IsNull(rsEM("ED_ORG")) Then
            locUnion = rsEM("ED_ORG")
        End If
        If Not IsNull(rsEM("ED_SIN")) Then
            locSIN = rsEM("ED_SIN")
        End If
        If Not IsNull(rsEM("ED_PAYROLL_ID")) Then
            locPayrollID = rsEM("ED_PAYROLL_ID")
        End If
        If Not IsNull(rsEM("ED_BENEFIT_GROUP")) Then locBenGroup = rsEM("ED_BENEFIT_GROUP") 'Ticket #24176 Franks 08/07/2013
        xEmpName = rsEM("ED_SURNAME") & ", " & rsEM("ED_FNAME")
        rsEM.Close
    End If
    'chkTerms ----------------------- end
    
    glbChgBenTermDate = ""
    ''Logic:
    ''1. Transfer out, don't do any thing for Manulife
    ''2. Termination, always terminate Benefits and Dependents
    'If Not glbTermTran Then 'Transfer Out
        ''glbChgBenTermDate = dlpBenCeaseDate
        'Ticket #24451 Franks 10/15/2013
        '"   For a MLF employee, when transferring out from Canada to another location, the benefits would have to end.
        'The benefit end date should equal the Transfer Out Date
        If Len(locCertNo) > 0 Then
            If IsOutFromBenGrpMatrix(xNewDiv) Then
                glbChgBenTermDate = xTermDate
            End If
        End If
    'End If
        
    Call WFC_NGS_Trans_TermTransferOut("Transfer Out", xTermDate, xNewDiv, xNewUnion, "")
        
    rsTB.Open "Term_HRSEQ", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If rsTB.EOF And rsTB.BOF Then
        glbTERM_Seq = 1
        rsTB.AddNew
    Else
        rsTB.MoveFirst 'Jaddy 10/28/99
        glbTERM_Seq = rsTB("TERM_SEQ_NEXT")
    End If
    rsTB("TERM_SEQ_NEXT") = glbTERM_Seq + 1
    rsTB.Update
    rsTB.Close
    
    Call WFCAUDITTermTransferOut(xTermDate, xNewDiv, xNewUnion, False)
    
    Call AUDIT_MANULIFE_TRANS_TermTransferOut(locCertNo)
    
    'Ticket #16395 - Pension System
    'Transfer Out - Trans Date, New Div
    toSOURCE = "IHR Transfer Out" 'Ticket #19954
    Call WFCPensionMasUpt(glbLEE_ID, "Transfer Out", xTermDate, xNewDiv, Year(CVDate(xTermDate)))
    'Ticket #19678 Franks 01/24/2011
    'On Transfer out:   (Plant Code equals "TILB" and Union Code is "C127") or (Plant Code equals "WHBY" and Union Code is "C222") the Hire Code equals "N".
    Call WFCHireCode(glbLEE_ID)

    'Ticket #21677 Franks 03/14/2012
    If xNewUnion = "NONE" Or xNewUnion = "EXEC" Then
        Call locWFCUpdPAMaster_TermTransferOut(glbLEE_ID, xTermDate, locSIN, locDiv, locPayrollID)
    End If
        
    'If Not modTermMove() Then Exit Sub
    'Call modTermMoveWFCTransferOut(xTERMDATE)
    Call modTermMoveWFCTransferOut(xTermDate, "TOUT", False)

    EID& = glbLEE_ID

    Call Term_Superv_General(EID&)

    Call Term_Reviewer_General(EID&)
    
    MDIMain.panHelp(0).FloodPercent = 100
    
    Call UpdEHScorrective(EID&, xEmpName)
    
    Call NukeEE2_General(EID&)
    
    '~~~~~~~~~~~~~~~~~~~~~~~~~'added by RAUBREY 5/23/97 ~~~~~~~~~~~~~~~~~~~~~~
    rsT_PARCO.Open "HRPARCO", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rsT_PARCO("PC_NUMBER_EMPLOYEES") = rsT_PARCO("PC_NUMBER_EMPLOYEES") - 1 'UPDATE FIELD WITH ACTUAL COUNT
    rsT_PARCO.Update
    rsT_PARCO.Close
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    'If Not glbTermTran Then 'Transfer Out
        SQLQ = "UPDATE Term_HREMP SET ED_OMERS=" & Date_SQL(xTermDate)
        SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
        gdbAdoIhr001X.Execute SQLQ
        Call UptLUserLDateLTime(glbTERM_Seq) 'Ticket #24355 Franks 09/17/2013
    'End If

    If glbAdv Then 'Ticket #15074
        Call Employee_Master_Integration(glbLEE_ID, "T" & Trim(Str(glbTERM_Seq)))
    End If

End Sub

Private Function WFCAUDITTermTransferOut(xTermDate, xNewDiv, xNewUnion, xIsTerm As Boolean)
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsTACheck As New ADODB.Recordset
Dim xPT As String, xDIV As String, XSNAME As String, xFName As String, xEmpType As String, xDOH As String, xSENDTE As String
Dim SQLQ As String, strFields As String
Dim xAdminBy As String
Dim xOldUnion, xReason
Dim xBatchID

On Error GoTo AUDIT_ERR

WFCAUDITTermTransferOut = False


If xIsTerm Then
    xReason = "EEIN"
    Call AuditFutureDataDele(glbLEE_ID, xTermDate) 'Ticket #24859 Franks 01/15/2014
Else
    xReason = "TOUT"
End If

glbChgTermReason = xReason
glbChgTermDate = xTermDate
    
rsTB.Open "SELECT ED_PT,ED_DIV,ED_SURNAME,ED_FNAME,ED_EMPTYPE,ED_DOH,ED_SENDTE,ED_ADMINBY,ED_ORG FROM HREMP WHERE ED_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If Not IsNull(rsTB("ED_PT")) Then   'Hemu - Gives an error when it's Null and this checking is not done
        xPT = rsTB("ED_PT")
    Else
        xPT = ""
    End If
    
    If Not IsNull(rsTB("ED_DIV")) Then 'George Apr 4,2006
        'xDiv = rsTB("ED_DIV")
        If IsNull(rsTB("ED_DIV")) Then xDIV = "" Else xDIV = rsTB("ED_DIV")
    Else
        xDIV = ""
    End If
    
    'Ticket #20884 Franks 10/20/2011
    If Not IsNull(rsTB("ED_ADMINBY")) Then
        If IsNull(rsTB("ED_ADMINBY")) Then xAdminBy = "" Else xAdminBy = rsTB("ED_ADMINBY")
    Else
        xAdminBy = ""
    End If
    
    XSNAME = rsTB("ED_SURNAME")
    xFName = rsTB("ED_FNAME")
    If IsNull(rsTB("ED_EMPTYPE")) Then
        xEmpType = ""
    Else
        xEmpType = rsTB("ED_EMPTYPE")
    End If
    xDOH = rsTB("ED_DOH")
    If IsNull(rsTB("ED_SENDTE")) Then
        xSENDTE = ""
    Else
        xSENDTE = rsTB("ED_SENDTE")
    End If
    If IsNull(rsTB("ED_ORG")) Then xOldUnion = "" Else xOldUnion = rsTB("ED_ORG")
Else
    xPT = ""
    xDIV = ""
    XSNAME = ""
    xFName = ""
    xEmpType = ""
    xDOH = ""
    xSENDTE = ""
    xAdminBy = ""
    xOldUnion = ""
End If
rsTB.Close
'Linamar doesn't need Audit records when Transfer Out
'WFC need Audit records when Transfer Out
'Ticket# 7337 For Linamar Interface
'If glbTermTran Or Not glbLinamar Then
    'strFields added by Bryan 02/Dec/05 Ticket#9899
    strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, "
    strFields = strFields & "AU_TREAS_TABL, AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_EMPTYPE, AU_SURNAME, "
    strFields = strFields & "AU_FNAME, AU_DOT, AU_TREAS, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_PAYROLL_ID,AU_VADIM2,AU_SIN,AU_SSN,AU_ADMINBY "
    rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    rsTA.AddNew
    rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
    rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
    rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
    rsTA("AU_NEWEMP") = "N"
    rsTA("AU_PTUPL") = xPT
    rsTA("AU_DIVUPL") = xDIV
    If glbSamuel Then 'Ticket #20884 Franks 10/20/2011
        rsTA("AU_ADMINBY") = xAdminBy
    End If
    rsTA("AU_EMPTYPE") = xEmpType
    rsTA("AU_SURNAME") = XSNAME
    rsTA("AU_FNAME") = xFName
    rsTA("AU_DOT") = xTermDate
    rsTA("AU_TREAS") = xReason 'clpCode(1)
    rsTA("AU_COMPNO") = "001"
    rsTA("AU_EMPNBR") = glbLEE_ID
    rsTA("AU_LDATE") = Date
    rsTA("AU_LUSER") = glbUserID
    rsTA("AU_LTIME") = Time$
    rsTA("AU_UPLOAD") = "N"
    rsTA("AU_TYPE") = "T"
    'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    'Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID,ED_SIN,ED_SSN FROM HREMP WHERE ED_EMPNBR =" & glbLEE_ID
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
        'Ticket #16749
        If glbWFC Then
            rsTA("AU_SIN") = rsEmp("ED_SIN")
            rsTA("AU_SSN") = rsEmp("ED_SSN")
        End If
    End If
    rsEmp.Close
    'End If
    rsTA.Update
    rsTA.Close
'End If

If glbLinamar Or glbWFC Or glbSamuel Then 'For Samuel Ticket #20884 Franks 10/20/2011
    Dim xKey, xCURRENTDIV, xJob
    xKey = "T" & glbTERM_Seq
    rsTB.Open "SELECT JH_JOB FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
    If Not rsTB.EOF Then
        xJob = rsTB!JH_JOB
    Else
        xJob = ""
    End If
    rsTA.Open "LN_TRALOG", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    rsTA.AddNew
    
    rsTA!TL_COMPNO = "001"
    rsTA!TL_EMPNBR = glbLEE_ID
    rsTA!TL_SURNAME = XSNAME
    rsTA!TL_FNAME = xFName
    If IsDate(xDOH) Then rsTA!TL_DOH = xDOH
    rsTA!TL_JOB = xJob
    'If glbTermTran Then
    If xIsTerm Then
        rsTA!TL_TYPE = "TERM"
        rsTA!TL_TCOMPLETE = "Y"
        xCURRENTDIV = xDIV
        rsTA!TL_NEWDIV = xCURRENTDIV
        rsTA!TL_NEWEMPNBR = glbLEE_ID
        If Len(xSENDTE) > 0 Then
            rsTA!TL_NEWDIVEDATE = xSENDTE
        End If
    Else
        'transfer out - begin
        rsTA!TL_TYPE = "TOUT"
        If glbWFC Then
            xCURRENTDIV = xNewDiv
            rsTA!TL_NEWDIV = xCURRENTDIV
        End If
        'Ticket #21677 Franks 03/14/2012 - union transfer
        If glbWFC Then
            If Len(xNewUnion) > 0 Then
                rsTA!TL_OLD_ORG = xOldUnion 'Trim(Mid(lblCurUnion.Caption, InStr(lblCurUnion.Caption, ":") + 2, 4))
                If Len(xOldUnion) > 0 Then
                    rsTA!TL_OLD_ORG_DESC = GetTABLDesc("EDOR", xOldUnion)
                End If
                rsTA!TL_NEW_ORG = xNewUnion
                rsTA!TL_NEW_ORG_DESC = GetTABLDesc("EDOR", xNewUnion)
            End If
        End If
        
        rsTA!TL_NEWEMPNBR = glbLEE_ID 'fglbEMPNBR
        rsTA!TL_NEWDIVEDATE = xTermDate
        rsTA!TL_TCOMPLETE = "N"
    End If 'transfer out end ------------------
    
    rsTA!TL_OLDDIV = xDIV
    rsTA!TL_OLDEMPNBR = glbLEE_ID
    If Len(xSENDTE) > 0 Then
        rsTA!TL_OLDDIVEDATE = xSENDTE
    End If
    rsTA!TL_TOREASON_TABL = "TERM"
    rsTA!TL_TIREASON_TABL = "SDJC"
    rsTA!TL_TOREASON = xReason ' clpCode(1)
    rsTA!TL_TERM_SEQ = glbTERM_Seq
    
    rsTA!TL_KEY = xKey
    rsTA!TL_CURRENTDIV = xCURRENTDIV
    
    rsTA("TL_LDATE") = Format(Now, "SHORT DATE")
    rsTA("TL_LUSER") = glbUserID
    rsTA("TL_LTIME") = Time$
    rsTA.Update
    rsTA.Close
    rsTA.Open "SELECT TL_KEY,TL_CURRENTDIV FROM LN_TRALOG WHERE TL_KEY='E" & glbLEE_ID & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do Until rsTA.EOF
        rsTA!TL_KEY = xKey
        rsTA!TL_CURRENTDIV = xCURRENTDIV
        rsTA.Update
        rsTA.MoveNext
    Loop
    rsTA.Close
'    gdbAdoIhr001.Execute "UPDATE LN_TRALOG SET TL_KEY='" & xKEY & "' WHERE TL_KEY='E" & glbLEE_ID & "'"

End If

WFCAUDITTermTransferOut = True

Exit Function

AUDIT_ERR:
glbFrmCaption$ = "Transfer In"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "UPDATE")
Call RollBack '29July99 js
Resume Next
End Function


Public Sub WFC_NGS_Trans_TermTransferOut(xType, xTermDate, xNewDiv, xNewUnion, xdlpDOther2) '#19266
Dim rsEmpee As New ADODB.Recordset
Dim rsEmpOther As New ADODB.Recordset
Dim SQLQ As String
Dim xUnion As String
Dim xSalHly As String
Dim xInSubGrp As String
Dim xLDate
Dim xNGSStart
Dim xCurPlant, xToPlant 'Ticket #23501 Franks 04/02/2013
    
    SQLQ = "SELECT ED_EMPNBR, ED_DIV, ED_ORG, ED_VADIM1, ED_VADIM2 FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID & " "
    rsEmpee.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmpee.EOF Then
        Exit Sub
    Else
        If IsNull(rsEmpee("ED_DIV")) Then glbEmpDiv = "" Else glbEmpDiv = rsEmpee("ED_DIV")
        If IsNull(rsEmpee("ED_ORG")) Then glbUNION = "" Else glbUNION = rsEmpee("ED_ORG")
        If IsNull(rsEmpee("ED_VADIM1")) Then glbWFCNGSSubGroup = "" Else glbWFCNGSSubGroup = rsEmpee("ED_VADIM1")
        If IsNull(rsEmpee("ED_VADIM2")) Then glbWFCPayGroup = "" Else glbWFCPayGroup = rsEmpee("ED_VADIM2")
    End If
    rsEmpee.Close
    
    'No NGS Sub Group, skip
    If Len(glbWFCNGSSubGroup) = 0 Then Exit Sub

    'Ticket #20385 Franks 05/31/2011
    'xLDate = dlpTermDate.Text 'Date
    xLDate = Date
    
    xNGSStart = ""
    SQLQ = "SELECT ER_EMPNBR,ER_OTHERDATE1 FROM HREMP_OTHER WHERE ER_EMPNBR = " & glbLEE_ID & ""
    rsEmpOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpOther.EOF Then
        If IsDate(rsEmpOther("ER_OTHERDATE1")) Then
            xNGSStart = rsEmpOther("ER_OTHERDATE1")
        End If
    End If
    rsEmpOther.Close
    
    ''Ticket #20385 Franks 05/31/2011
    ''No NGS Effective Date, skip
    'If Len(xNGSStart) = 0 Then Exit Sub

    If glbUNION = "NONE" Or glbUNION = "EXEC" Then
        xSalHly = "Y"
    Else
        xSalHly = "N"
    End If

    If xType = "Transfer Out" Then
        If Len(xNewUnion) > 0 Then 'Ticket #21677 Franks 03/14/2012
            xInSubGrp = getNGSSubGrpFromMatrix(xNewDiv, xNewUnion)
        Else
            xInSubGrp = getNGSSubGrpFromMatrix(xNewDiv, glbUNION)
        End If
        If Len(xInSubGrp) = 0 Then
            'Ticket #23501 Franks 04/02/2013
            '"   If the Plant in the Transfer To Division equals the Plant from the employee's record
            'do not create the NGS Audit record and do not populate the NGS End Date field.
            xCurPlant = getSectionByDiv(glbEmpDiv)
            xToPlant = getSectionByDiv(xNewDiv)
            If xCurPlant = xToPlant Then
                'transfer between same plant, do not change NGS
            Else
                Call Upt_EmpOtherByField(glbLEE_ID, "ER_OTHERDATE2", CVDate(xTermDate))
                'Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", lStr("Other Date 2"), "", CVDate(dlpTermDate.Text), xLDate)
                'Ticket #22409 Franks 08/16/2012 send NGS End Date only when the employee transfer out NGS group
                Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", lStr("Other Date 2"), "", CVDate(xTermDate), xLDate)
            End If
        End If
        'Ticket #22409 Franks 08/16/2012 do not send NGS End Date between unions
        ''Ticket #21822 Franks 04/10/2012 - Send NGS End Date to NGS for transfer between unions or Div
        'Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", lStr("Other Date 2"), "", CVDate(dlpTermDate.Text), xLDate)
    End If
    If xType = "Termination" Then
        If IsDate(xdlpDOther2) Then
            Call Upt_EmpOtherByField(glbLEE_ID, "ER_OTHERDATE2", CVDate(xdlpDOther2))
            'Call NGSAuditAdd(glbLEE_ID, "M", "Termination", "Transfer Out Date", "", CVDate(dlpTermDate.Text), xLDate)
            'Call NGSAuditAdd(glbLEE_ID, "M", "Transfer Out", "To Division", glbEmpDiv, Trim(xNewDiv), xLDate)
            Call NGSAuditAdd(glbLEE_ID, "M", "Termination", lStr("Other Date 2"), "", CVDate(xdlpDOther2), xLDate)
        End If
    End If
End Sub

Public Function getSectionByDiv(xDIV) 'Ticket #23501 Franks 04/02/2013
Dim rsDiv As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDIV & "' "
    If rsDiv.State <> 0 Then rsDiv.Close
    rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsDiv.EOF Then
        If Not IsNull(rsDiv("DV_SECTION")) Then
            retVal = rsDiv("DV_SECTION")
        End If
    End If
    rsDiv.Close
    getSectionByDiv = retVal
End Function

Public Function IsOutFromBenGrpMatrix(xDIV)
Dim rsBenGrpMrx As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HR_BENEFITS_GROUP_MATRIX WHERE (1=1) " ' BM_BENEFIT_GROUP = '" & NewBGroup & "' "
    If Len(xDIV) > 0 Then
        SQLQ = SQLQ & "AND BM_DIV = '" & xDIV & "' "
    End If
    rsBenGrpMrx.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsBenGrpMrx.EOF Then
        'xCovClass = rsBenGrpMrx("BM_BENEFIT_CLASS")
        'xBenAccount = rsBenGrpMrx("BM_BENEFIT_ACCOUNT")
        retVal = True
    End If
    rsBenGrpMrx.Close
    IsOutFromBenGrpMatrix = retVal
End Function

Public Function AUDIT_MANULIFE_TRANS_TermTransferOut(xCertNo) 'No AU_CEASEDATE in HRAUDIT, Jerry said we will add it in next release
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim rsBene As New ADODB.Recordset
Dim rsDepend As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
Dim SQLQ As String
'''On Error GoTo AUDIT_ERR
AUDIT_MANULIFE_TRANS_TermTransferOut = False

'BENEFIT End Date
If Len(xCertNo) = 0 Or Len(glbChgBenTermDate) = 0 Then
    Exit Function
End If

rsTB.Open "SELECT ED_DIV, ED_SECTION, ED_USER_TEXT1,ED_USER_TEXT2,ED_USER_NUM1  FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID, gdbAdoIhr001, adOpenKeyset
If rsTB.EOF Then
    rsTB.Close:    GoTo MODNOUPD_Den
End If
If IsNull(rsTB("ED_USER_TEXT1")) Then 'Certificate #
    rsTB.Close:    GoTo MODNOUPD_Den
Else
    If Len(Trim(rsTB("ED_USER_TEXT1"))) = 0 Then
        rsTB.Close:    GoTo MODNOUPD_Den
    End If
End If

'Benefits
SQLQ = "SELECT * FROM HRBENFT WHERE NOT(BF_POLICY IS NULL) AND BF_EMPNBR = " & glbLEE_ID
rsBene.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsBene.EOF Then
    rsBene.Close
    GoTo MODNOUPD_Ben 'Exit Function
End If


Do While Not rsBene.EOF
    If Len(rsBene("BF_POLICY")) > 0 Then
        If Not IsDate(rsBene("BF_CEASEDATE")) Then 'No Benefit End Date
            If rsTA.State <> 0 Then rsTA.Close
            rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
            
            rsTA.AddNew
            rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
            rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
            rsTA("MT_PT_TABL") = "EDPT"
            rsTA("MT_TYPE") = "T"
            rsTA("MT_BENEFIT") = rsBene("BF_BCODE")
            rsTA("MT_EDATE") = rsBene("BF_EDATE")
            rsTA("MT_CEASEDATE") = glbChgBenTermDate
            rsTA("MT_COVER") = rsBene("BF_COVER")
            rsTA("MT_COMPNO") = "001"
            rsTA("MT_EMPNBR") = glbLEE_ID
            rsTA("MT_POLICY_NO") = rsBene("BF_POLICY")
            rsTA("MT_ACCOUNT_NO") = rsTB("ED_USER_NUM1")
            rsTA("MT_CERT_NO") = rsTB("ED_USER_TEXT1")
            rsTA("MT_COVERAGE_CLASS") = rsTB("ED_USER_TEXT2")
            rsTA("MT_UPLOAD") = "N"
            rsTA("MT_LUSER") = glbUserID
            If CVDate(glbChgBenTermDate) < CVDate(Date) Then 'Ticket #14867
                rsTA("MT_LDATE") = Date
            Else
                rsTA("MT_LDATE") = Format(glbChgBenTermDate, "SHORT DATE")
            End If
            rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
            rsTA("MT_LTIME") = Time$
            
            rsTA.Update
        End If
    End If
    rsBene.MoveNext
Loop
rsBene.Close

MODNOUPD_Ben:

SQLQ = "SELECT * FROM HRDEPEND WHERE DP_EMPNBR = " & glbLEE_ID
rsDepend.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rsDepend.EOF Then
    rsDepend.Close
    GoTo MODNOUPD_Den 'Exit Function
End If
    
Do While Not rsDepend.EOF
    If Not IsDate(rsDepend("DP_EDATE")) Then 'No Benefit End Date
        If rsTA.State <> 0 Then rsTA.Close
        rsTA.Open "SELECT * FROM HR_MANULIFE_TRAN_AUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        
        rsTA.AddNew
        rsTA("MT_LOC_TABL") = "EDLC": rsTA("MT_SECTION_TABL") = "EDSE": rsTA("MT_EMP_TABL") = "EDEM"
        rsTA("MT_ORG_TABL") = "EDOR": rsTA("MT_BENEFIT_TABL") = "BNCD"
        rsTA("MT_PT_TABL") = "EDPT"
        rsTA("MT_TYPE") = "T"
        rsTA("MT_DEPFNAME") = rsDepend("Dp_FName")
        rsTA("MT_DEPSNAME") = rsDepend("DP_SNAME")
        rsTA("MT_DEPSEX") = rsDepend("DP_SEX")
        rsTA("MT_DEPDOB") = rsDepend("DP_DOB")
        rsTA("MT_DEPRELATE") = rsDepend("DP_RELATE")
        rsTA("MT_DEPSMOKER") = rsDepend("DP_SMOKER")
        rsTA("MT_DEPSTATUS") = rsDepend("DP_STATUS")
        rsTA("MT_DEPSIN") = rsDepend("DP_SIN")
        rsTA("MT_DEPSDATE") = rsDepend("DP_SDATE")
        rsTA("MT_DEPEDATE") = glbChgBenTermDate
        rsTA("MT_DENTAL") = rsDepend("DP_DENTAL")
        rsTA("MT_MEDICAL") = rsDepend("DP_MEDICAL")
        rsTA("MT_OTHER") = rsDepend("DP_OTHER")
        rsTA("MT_COMPNO") = "001"
        rsTA("MT_EMPNBR") = glbLEE_ID
        rsTA("MT_ACCOUNT_NO") = rsTB("ED_USER_NUM1")
        rsTA("MT_CERT_NO") = rsTB("ED_USER_TEXT1")
        rsTA("MT_COVERAGE_CLASS") = rsTB("ED_USER_TEXT2")
        rsTA("MT_UPLOAD") = "N"
        rsTA("MT_LUSER") = glbUserID
        If CVDate(glbChgBenTermDate) < CVDate(Date) Then 'Ticket #14867
            rsTA("MT_LDATE") = Date
        Else
            rsTA("MT_LDATE") = Format(glbChgBenTermDate, "SHORT DATE")
        End If
        rsTA("MT_TRAN_DATE") = Format(Date, "SHORT DATE")
        rsTA("MT_LTIME") = Time$
        rsTA.Update
    End If
    rsDepend.MoveNext
Loop
rsDepend.Close

MODNOUPD_Den:

AUDIT_MANULIFE_TRANS_TermTransferOut = True
Exit Function
AUDIT_ERR:

glbFrmCaption$ = "Transfer Out"
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING MANULIFE AUDIT RECORD", "MANULIFE AUDIT FILE", "UPDATE")
'If gintRollBack% = False Then Resume Next Else Unload Me
Resume Next

End Function

Public Sub WFCHireCode(xEmpNo)
Dim rsLEmp As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT ED_EMPNBR, ED_EMPTYPE, ED_SECTION, ED_ORG, ED_HIRECODE FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND ED_EMPTYPE = 'Y' "
    rsLEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsLEmp.EOF Then
        If Not IsNull(rsLEmp("ED_SECTION")) And Not IsNull(rsLEmp("ED_ORG")) Then
            If (rsLEmp("ED_SECTION") = "TILB" And rsLEmp("ED_ORG") = "C127") Or (rsLEmp("ED_SECTION") = "WHBY" And rsLEmp("ED_ORG") = "C222") Then
                rsLEmp("ED_HIRECODE") = "N"
                rsLEmp.Update
            End If
        End If
    End If
    rsLEmp.Close
End Sub

Public Sub locWFCUpdPAMaster_TermTransferOut(xEmpNo, xTranDate, xSIN, xCurDiv, locPayrollID)
'o   If the Union Code changes to "NONE" or "EXEC", a PA Master must be created
'"   Earned Pension is calculated using the Hourly Year End Pension & PA Update rules. (Credited Service months * Benefit Rate)
'Frank Note: the Union Transfer Out create the Pension Master with status code X first
'so this PA Master update function will reuse the same Earning Pension from Pension Master
Dim rsPen As New ADODB.Recordset
Dim rsPAMaster As New ADODB.Recordset
Dim SQLQ As String
Dim xYear
Dim xEarnPen, xTotal
    If Len(xSIN) = 0 Then Exit Sub
    If Not IsDate(xTranDate) Then Exit Sub
    
    xYear = Year(CVDate(xTranDate))
    SQLQ = "SELECT * FROM HRP_PENSION_MASTER WHERE PE_SIN = '" & xSIN & "' "
    SQLQ = SQLQ & "AND LEFT(PE_PENSIONTYPE,2) = 'DB' "
    SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
    SQLQ = SQLQ & "AND PE_DB_STATUS = 'X' "
    SQLQ = SQLQ & "AND PE_HRLYSAL = 'Hourly' "
    SQLQ = SQLQ & "ORDER BY PE_YEAR_DATE DESC"
    rsPen.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsPen.EOF Then
        xEarnPen = 0
        If Not IsNull(rsPen("PE_CREDITED_SERV")) And Not IsNull(rsPen("PE_BENEFIT_RATE")) Then
            xEarnPen = rsPen("PE_CREDITED_SERV") * rsPen("PE_BENEFIT_RATE")
        End If
        SQLQ = "SELECT * FROM HRP_PA_MASTER WHERE PE_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND PE_YEAR_DATE = " & xYear & " "
        SQLQ = SQLQ & "AND PE_HRLYSAL = 'Hourly' "
        rsPAMaster.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsPAMaster.EOF Then
            rsPAMaster.AddNew
            rsPAMaster("PE_COUNTRY") = rsPen("PE_COUNTRY")
            rsPAMaster("PE_SIN") = rsPen("PE_SIN")
            rsPAMaster("PE_EMPNBR") = rsPen("PE_EMPNBR")
        End If
        rsPAMaster("PE_SURNAME") = rsPen("PE_SURNAME")
        rsPAMaster("PE_FNAME") = rsPen("PE_FNAME")
        rsPAMaster("PE_DIV") = xCurDiv  'Left(lblCurDiv.Caption, 4) ' rsPen("PE_DIV")
        rsPAMaster("PE_SECTION") = rsPen("PE_SECTION")
        rsPAMaster("PE_YEAR_DATE") = xYear
        rsPAMaster("PE_HRLYSAL") = rsPen("PE_HRLYSAL")
        If Len(locPayrollID) > 0 Then
            rsPAMaster("PE_PAYROLL_ID") = locPayrollID
        End If
        rsPAMaster("PE_DMD_PENEARN") = xEarnPen
        '"   (Earned Pension * 9) - 600
        xTotal = (xEarnPen * 9) - 600
        If xTotal < 0 Then xTotal = 0
        rsPAMaster("PE_TOTAL_DBPA") = xTotal
        rsPAMaster("PE_TOTAL_PA") = xTotal
        rsPAMaster("PE_LDATE") = Date
        rsPAMaster("PE_LTIME") = Time$
        rsPAMaster("PE_LUSER") = glbUserID
        rsPAMaster.Update
    End If
    
End Sub

Private Function modTermMoveWFCTransferOut(xTermDate, xTermReason As String, xIsTerm As Boolean)
Dim X%
Dim EEID&, TReason$, DtTm  As Variant, TRDesc$
Dim TComment$
Dim TRehire$
Dim TCause

Screen.MousePointer = HOURGLASS

modTermMoveWFCTransferOut = False
DtTm = xTermDate 'dlpTermDate.Text
EEID& = glbLEE_ID ' lblEEID
If xIsTerm Then
    TReason$ = xTermReason 'clpCode(1).Text
    TRDesc$ = getCodeDesc("TERM", xTermReason) '"" '"Transfer Out"
Else
    TReason$ = "TOUT"  'clpCode(1).Text
    TRDesc$ = "Transfer Out" 'clpCode(1).Caption
End If
TComment$ = "" 'txtComments
TRehire$ = "Yes" 'lblRehire
TCause = "" 'clpCode(2).Text

gdbAdoIhr001.BeginTrans
'gdbAdoIhr001X.BeginTrans

X% = TERM_LIST(EEID&, DtTm, TReason$, TRDesc$, TComment$, TRehire$, TCause)
MDIMain.panHelp(0).FloodPercent = 5
X% = TERM_BASIC(EEID&)
MDIMain.panHelp(0).FloodPercent = 10
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EDUCSEM(EEID&)                  'laura nov 5, 1997
MDIMain.panHelp(0).FloodPercent = 13      '
If Not X Then GoTo modTermMoveErr_Msg    '
X% = TERM_ATTENDANCE(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 15
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_ATTENDANCE_HISTORY(EEID&, DtTm)
MDIMain.panHelp(0).FloodPercent = 18
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_JOB(EEID&)
MDIMain.panHelp(0).FloodPercent = 20
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_PERFORM(EEID&)
MDIMain.panHelp(0).FloodPercent = 22
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_SALARY(EEID&)
MDIMain.panHelp(0).FloodPercent = 25
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HealthSafety(EEID&)
MDIMain.panHelp(0).FloodPercent = 28
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_BENEFITS(EEID&)
MDIMain.panHelp(0).FloodPercent = 30
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_DEPEND(EEID&)
MDIMain.panHelp(0).FloodPercent = 31
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HealthCost(EEID&)
MDIMain.panHelp(0).FloodPercent = 32
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_OHS_Contact(EEID&)
MDIMain.panHelp(0).FloodPercent = 35
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_COMMENTS(EEID&)               'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 38
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_COBRA(EEID&)
MDIMain.panHelp(0).FloodPercent = 39
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_OHS_Corrective(EEID&)
MDIMain.panHelp(0).FloodPercent = 40
If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_ROOT_CAUSES(EEID&)

If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_CLAIM_MEDICAL(EEID&)
If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_FORM7_SECTIONS(EEID&)

'Ticket #21463
If Not X Then GoTo modTermMoveErr_Msg
X% = Term_OHS_FORM9(EEID&)

MDIMain.panHelp(0).FloodPercent = 43
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_DOLENT(EEID&)

'Ticket #28789 - Actual Amounts Details
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_DOLENT_ACTDTL(EEID&)

MDIMain.panHelp(0).FloodPercent = 45
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_ENTHRS(EEID&)                 'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 46
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EARN(EEID&)                   'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 48
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EDU(EEID&)                    'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 50
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMPSKL(EEID&)                 'FRANK 4/5/2000
MDIMain.panHelp(0).FloodPercent = 52
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_TRADE(EEID&)                  'FRANK 4/5/2000
If Not X Then GoTo modTermMoveErr_Msg
MDIMain.panHelp(0).FloodPercent = 53
X% = TERM_COUNSEL(EEID&)                ' dkostka - 10/02/2001
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HREMPHIS(EEID&)                ' Hemu - 06/30/2004
If Not X Then GoTo modTermMoveErr_Msg
'If glbWFC Then
X% = TERM_EMPOTHER(EEID&)                  'FRANK 11/05/2004
If Not X Then GoTo modTermMoveErr_Msg
'End If
X% = TERM_USERDEFINE_TABLE(EEID&)          'Hemu - 02/28/2008
If Not X Then GoTo modTermMoveErr_Msg

X% = TERM_SUCCESSION(EEID&)          'George 04/04/2006 #10595
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_LANGUAGE(EEID&)          'George 04/04/2006 #10595
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMP_FLAGS(EEID&)          'Bryan 05/04/2006
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_GLDIST(EEID&)             'Bryan 05/04/2006
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMPADP(EEID&)                  'FRANK 06/08/2006
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_EMPPAYROLL_TRANSACTION(EEID&)  'FRANK 03/18/2010 Ticket #18232
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_FOLLOW_UP(EEID&)  'Hemu 08/27/2010 Ticket #18668
If Not X Then GoTo modTermMoveErr_Msg
X% = TERM_HREEO(EEID&)  'Ticket #25669 Franks 06/24/2014
If Not X Then GoTo modTermMoveErr_Msg

If gsAttachment_DB Then
    X% = TERM_HRDOC_EMP(EEID&)                  'FRANK 01/10/2006
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_JOB_HISTORY(EEID&)          'George 01/19/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_COMMENTS(EEID&)          'George 01/26/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_HEALTH_SAFETY(EEID&)          'George 02/17/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_HEALTH_SAFETY_2(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    
    If glbWSIBModule Then
        X% = TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7(EEID&)
        If Not X Then GoTo modTermMoveErr_Msg
        X% = TERM_HRDOC_OHS_WRITTEN_OFFER(EEID&)
        If Not X Then GoTo modTermMoveErr_Msg
    End If
    
    X% = TERM_HRDOC_COUNSEL(EEID&)          'George 01/26/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_PERFORM_HISTORY(EEID&)          'George 01/26/2006 #10266
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_EDSEM(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_EDSEM_RETEST(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_HREDU(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg '
    X% = TERM_HRDOC_HRDOLENT(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_TRADE(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_ATTENDANCE(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
    X% = TERM_HRDOC_EMP_FLAGS(EEID&)
    If Not X Then GoTo modTermMoveErr_Msg
End If '


MDIMain.panHelp(0).FloodPercent = 55

If Not X Then GoTo modTermMoveErr_Msg

X% = InputHREMPEQU_DOT_WFCTran(EEID&)

gdbAdoIhr001.CommitTrans
'gdbAdoIhr001X.CommitTrans

modTermMoveWFCTransferOut = True

Screen.MousePointer = DEFAULT

Exit Function

modTermMoveErr_Msg:
Screen.MousePointer = DEFAULT

MsgBox ("Problem Creating Audit record - Transfer Out Aborted")

End Function

Private Function InputHREMPEQU_DOT_WFCTran(EmpN As Long)
Dim SQLQ As String
Dim dynEmp As New ADODB.Recordset

SQLQ = "SELECT * FROM HREMPEQU WHERE HREMPEQU.EQ_EMPNBR = "
SQLQ = SQLQ & EmpN


dynEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If dynEmp.RecordCount > 0 Then
    'Release 8.0 - Ticket #24309: Addition option to enter Terminated Employees on hte Employment Equity Survey screen
    'SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = " & Date_SQL(dlpTermDate.Text) & " "
    SQLQ = "UPDATE HREMPEQU SET HREMPEQU.EQ_DOT = " & Date_SQL(dlpTermDate.Text) & ", EQ_TYPE = 'T' "
    SQLQ = SQLQ & "WHERE HREMPEQU.EQ_EMPNBR = " & EmpN
    gdbAdoIhr001.Execute SQLQ
End If

End Function

Public Function Term_Superv_General(xEmpNo)
Term_Superv_General = False
Dim SQLQDel As String, SQLQCom As String, strTable As String
Dim dynHRAT As New ADODB.Recordset
Dim strComm

On Error GoTo Database_Err

Screen.MousePointer = HOURGLASS

SQLQCom = "UPDATE HR_ATTENDANCE SET AD_SUPER = NULL WHERE AD_SUPER = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_ATTENDANCE_HISTORY SET AH_SUPER = NULL WHERE AH_SUPER = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_PERFORM_HISTORY SET PH_REPTAU = NULL WHERE PH_REPTAU = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU = NULL WHERE JH_REPTAU = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom
SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU2 = NULL WHERE JH_REPTAU2 = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom
SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU3 = NULL WHERE JH_REPTAU3 = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom
SQLQCom = "UPDATE HR_JOB_HISTORY SET JH_REPTAU4 = NULL WHERE JH_REPTAU4 = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_EMPNOT = NULL WHERE EC_EMPNOT = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_SUPERVISOR = NULL WHERE EC_SUPERVISOR = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

SQLQCom = "UPDATE HR_COUNSEL SET CL_COUBY = NULL WHERE CL_COUBY = " & xEmpNo
gdbAdoIhr001.Execute SQLQCom

Screen.MousePointer = DEFAULT

Term_Superv_General = True
Exit Function

Database_Err:
glbFrmCaption$ = "Termination/Transfer Out"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Superv_General", strTable, "TERMINATE")

End Function

Public Function Term_Reviewer_General(xEmpNo)
Dim SQLQDel As String, SQLQCom As String, strTable As String
Dim dynHRAT As New ADODB.Recordset
Dim strComm

On Error GoTo Database_Err
Term_Reviewer_General = False

strTable = "HR_SUCCESSION"
SQLQCom = "SELECT EU_REVIEWER FROM HR_SUCCESSION WHERE EU_REVIEWER = " & xEmpNo

dynHRAT.Open SQLQCom, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

Screen.MousePointer = HOURGLASS

If dynHRAT.RecordCount >= 1 Then
    dynHRAT.MoveFirst
    While Not dynHRAT.EOF
        dynHRAT("EU_REVIEWER") = 0
        dynHRAT.Update
        dynHRAT.MoveNext
    Wend
End If
dynHRAT.Close

Screen.MousePointer = DEFAULT

Term_Reviewer_General = True
Exit Function

Database_Err:
glbFrmCaption$ = "Termination/Transfer Out"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Reviewer_General", strTable, "TERMINATE")

End Function

Public Sub UpdEHScorrective(xEmpNo, xName)
    Dim SQLQ
    If InStr(1, xName, "'") > 0 Then
        xName = Replace(xName, "'", "''")
        If glbLinamar Then
            SQLQ = "update HR_OHS_CORRECTIVE set CR_TERM_EMPNAME ='" & xName & "' WHERE CR_ASSIGNED =" & xEmpNo
        Else
            SQLQ = "update HR_OHS_CORRECTIVE set CR_ASSIGNED  = Null,CR_TERM_EMPNAME ='" & xName & "' WHERE CR_ASSIGNED =" & xEmpNo
        End If
    Else
        If glbLinamar Then
            SQLQ = "update HR_OHS_CORRECTIVE set CR_TERM_EMPNAME ='" & xName & "' WHERE CR_ASSIGNED =" & xEmpNo
        Else
            SQLQ = "update HR_OHS_CORRECTIVE set CR_ASSIGNED  = Null,CR_TERM_EMPNAME ='" & xName & "' WHERE CR_ASSIGNED =" & xEmpNo
        End If
    End If
    gdbAdoIhr001.Execute SQLQ
End Sub

Public Sub NukeEE2_General(EEID As Long)
Dim snapEETables As New ADODB.Recordset
Dim SQLQ As String, TabName$
Dim EEIDAlias$

On Error GoTo NukeEE2_General_Err
Dim rsSE As New ADODB.Recordset
Dim xUserID As String
rsSE.Open "SELECT USERID FROM HR_SECURE_BASIC WHERE EMPNBR=" & EEID&, gdbAdoIhr001, adOpenStatic
If Not rsSE.EOF Then
    xUserID = rsSE("USERID")
    Call NukeUSERID(xUserID)
End If
rsSE.Close

SQLQ = "SELECT * FROM INFO_HR_TABLES "
SQLQ = SQLQ & " WHERE Employee_Keyed <>0"
SQLQ = SQLQ & " AND TERMINATION_TABLE=0"
'Ticket #22367, Ticket #20367 - Do not delete employee photo
SQLQ = SQLQ & " AND Table_Name <>'HR_PHOTO'"
'Ticket #25669 Franks 06/24/2014 - comment out the following code
'SQLQ = SQLQ & " AND Table_Name <>'HREEO'"

'Ticket #20415 - Add Serial # to the select statement so custom tables also gets employee # changed.
'Serial 9999 is by default for all standard info:HR table.
'SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL = '" & glbCompSerial & "')"
'Ticket #20893 Franks 09/02/2011 - only remove data for the standard INFO:HR tables
SQLQ = SQLQ & " AND (SERIAL = 'S/N - 9999W' OR SERIAL IS NULL)"

snapEETables.Open SQLQ, gdbAdoIhr001, adOpenStatic

If snapEETables.RecordCount < 1 Then Exit Sub
snapEETables.MoveFirst

While Not snapEETables.EOF
    TabName$ = snapEETables("Table_Name")
    If UCase(Right(TabName$, 3)) <> "WRK" Then
        EEIDAlias$ = snapEETables("EMPNBR_Alias")
        If glbVadim And TabName$ = "HRBENFT" Then ' special process for vadim integration
            gdbAdoIhr001.BeginTrans
            gdbAdoIhr001.Execute "UPDATE HRBENFT SET BF_LUSER='VADIM_INTEGRATION' WHERE BF_EMPNBR=" & EEID&
            gdbAdoIhr001.CommitTrans
        End If
        Call NukeEERows2_General(TabName$, EEIDAlias$, EEID&)
    ElseIf UCase(TabName$) = "HR_VACTIMEOFF_REQ_WRK" Then     'Ticket #25459
        EEIDAlias$ = snapEETables("EMPNBR_Alias")
        If glbVadim And TabName$ = "HRBENFT" Then ' special process for vadim integration
            gdbAdoIhr001.BeginTrans
            gdbAdoIhr001.Execute "UPDATE HRBENFT SET BF_LUSER='VADIM_INTEGRATION' WHERE BF_EMPNBR=" & EEID&
            gdbAdoIhr001.CommitTrans
        End If
        Call NukeEERows2_General(TabName$, EEIDAlias$, EEID&)
    End If
    snapEETables.MoveNext
Wend
If glbAxxent Then
    TabName$ = "HRRSP"
    EEIDAlias$ = "RS_EMPNBR"
    Call NukeEERows2_General(TabName$, EEIDAlias$, EEID&)
End If

snapEETables.Close

Call UpdVacTimeRequest(EEID&, "D")

If glbCompSerial = "S/N - 2362W" Then 'CITY OF SARNIA
    SQLQ = "DELETE FROM HR_OHS_REOCCURENCE WHERE CC_EMPNBR =" & EEID & " "
    gdbAdoIhr001.Execute SQLQ
End If

If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
    SQLQ = "DELETE FROM HR_PERFORM_FRIESEN WHERE PH_EMPNBR =" & EEID & " "
    gdbAdoIhr001.Execute SQLQ
End If

SQLQ = "DELETE FROM HR_PAYROLL_TRANSACTION WHERE PT_EMPNBR =" & EEID & " "
gdbAdoIhr001.Execute SQLQ
    
Exit Sub

NukeEE2_General_Err:
glbFrmCaption$ = "Delete Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_TABLES Error", "TabName$", "Search")
Call RollBack '29July99 js

End Sub

Public Sub NukeEERows2_General(TabName As String, EEIDAlias As String, EEID As Long)
' returns number of records found for ee in table
Dim Rows%, SQLQ As String
Dim gdbESS As New ADODB.Connection

On Error GoTo NukeEERows2_General_Err

If TabName$ = "HREMPEQU" Then
    Exit Sub
End If

If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
    If gdbESS <> "" Then
        gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
    End If
End If

SQLQ = "DELETE FROM " & TabName
SQLQ = SQLQ & " WHERE " & EEIDAlias & " = " & EEID

If Not glbSQL And Not glbOracle And (TabName = "HR_TIMESHEET" Or TabName = "HR_TIMESHEET_MODS" Or TabName = "HR_VACTIMEOFF_REQ" Or TabName = "HR_VACTIMEOFF_REQ_ARCHIVE") Then
    If gdbESS <> "" Then
        gdbESS.Execute SQLQ
    End If
Else
    gdbAdoIhr001.Execute SQLQ
End If

Exit Sub

NukeEERows2_General_Err:
glbFrmCaption$ = "Nuke Rows"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Delete EE Rows", TabName$, "Delete")
Call RollBack '29July99 js

End Sub

Public Sub UptLUserLDateLTime(xTERM_Seq) 'Ticket #24355 Franks 09/17/2013
Dim SQLQ As String
    'update Term_HREMP
    SQLQ = "UPDATE Term_HREMP SET ED_LDATE = " & Date_SQL(Date) & ", "
    SQLQ = SQLQ & "ED_LUSER = '" & glbUserID & "', "
    SQLQ = SQLQ & "ED_LTIME = '" & Time$ & "' "
    SQLQ = SQLQ & "WHERE TERM_SEQ=" & xTERM_Seq
    gdbAdoIhr001X.Execute SQLQ
    
    'update Term_HRTRMEMP
    SQLQ = "UPDATE Term_HRTRMEMP SET Term_LDATE = " & Date_SQL(Date) & ", "
    SQLQ = SQLQ & "Term_LUSER = '" & glbUserID & "', "
    SQLQ = SQLQ & "Term_LTIME = '" & Time$ & "' "
    SQLQ = SQLQ & "WHERE TERM_SEQ=" & xTERM_Seq
    gdbAdoIhr001X.Execute SQLQ
End Sub

Public Function WFCNamesMatched(xCandidate, xLEE_SName, xLEE_FName)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ
Dim retVal As Boolean
    retVal = False
    SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_CANDIDATE = " & xCandidate & " "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.BOF Then
        If Not IsNull(rsTemp("SF_SURNAME")) And Not IsNull(rsTemp("SF_FNAME")) Then
            If Trim(rsTemp("SF_SURNAME")) = Trim(xLEE_SName) And Trim(rsTemp("SF_FNAME")) = Trim(xLEE_FName) Then
                retVal = True
            End If
        End If
    End If
    rsTemp.Close
    WFCNamesMatched = retVal
End Function

Public Sub WFCHRSoftProcUpt(xFormName, Optional xEmpNo = 0)
Dim SQLQ As String
If glbCandidate > 0 Then
    If xFormName = "frmEEBASIC" Then
        SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_DEMO = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
        SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
        gdbAdoIhr001.Execute SQLQ
    End If
    If xFormName = "frmEESTATS" Then
        SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_STATUS = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
        SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
        gdbAdoIhr001.Execute SQLQ
    End If
    If xFormName = "frmEPOSITION" Then
        If IsWFCHRSFUptSuccess("frmEPOSITION", glbCandidate) Then
            SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_POSITION = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
            SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    If xFormName = "frmESALARY" Then
        If IsWFCHRSFUptSuccess("frmESALARY", glbCandidate) Then
            SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_SALARY = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
            SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
            gdbAdoIhr001.Execute SQLQ
            SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_PROCESSED = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
            SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    If xFormName = "frmEREHIRE" Then
        SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_REHIRE = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
        SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
        gdbAdoIhr001.Execute SQLQ
    End If
    If xFormName = "frmEWFCProm" Then
        SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_PROCESSED = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
        SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
        gdbAdoIhr001.Execute SQLQ
    End If
    If xFormName = "frmETRANIN" Then
        SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_PROCESSED = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
        SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
        gdbAdoIhr001.Execute SQLQ
        If xEmpNo > 0 Then
            SQLQ = "UPDATE HREMP SET ED_CANDIDATE = " & glbCandidate & " WHERE ED_EMPNBR = " & xEmpNo
            gdbAdoIhr001.Execute SQLQ
        End If
    End If
    If xFormName = "UNHIRES" Then
        SQLQ = "UPDATE HRSF_XML_IMPORT SET SF_UPT_PROCESSED = 1 WHERE SF_CANDIDATE = " & glbCandidate & " "
        SQLQ = SQLQ & "AND SF_ID = " & glbCand_SF_ID & " "
        gdbAdoIhr001.Execute SQLQ
    End If
End If
End Sub

Public Function getWFCRetireDate(xDATE) ' Ticket #24695 Franks 11/26/2013
'"   Regular New Hire fills in the Normal Retirement Date.
'Should be the first day of the month following the employee's 65th birthday.
Dim xYYYY, xMM, xDD
Dim retVal
    retVal = xDATE
    If IsDate(xDATE) Then
        retVal = DateAdd("yyyy", 65, CVDate(xDATE))
        retVal = DateAdd("m", 1, retVal)
        xYYYY = Year(retVal)
        xMM = MonthName(month(retVal))
        xDD = "1"
        retVal = CVDate(xMM & " 1," & xYYYY)
    Else
        retVal = Null
    End If
    getWFCRetireDate = retVal
End Function

Public Function getPosMasterValueByField(xJob, xFieldName) 'Ticket #24767 Franks 12/10/2013
Dim rsMaster As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xJob & "' "
    rsMaster.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsMaster.EOF Then
        If Not IsNull(rsMaster(xFieldName)) Then
            retVal = rsMaster(xFieldName)
        End If
    End If
    rsMaster.Close
    getPosMasterValueByField = retVal
End Function

Public Function getHRTABLCodeFromDesc(xName, xDesc)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    If Not IsNull(xDesc) Then
        SQLQ = "SELECT * FROM HRTABL WHERE TB_DESC = '" & Replace(xDesc, "'", "''") & "' "
        SQLQ = SQLQ & "AND TB_NAME = '" & xName & "' "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsTemp.EOF Then
            SQLQ = "SELECT * FROM HRTABL WHERE TB_DESC LIKE UPPER('%" & Replace(xDesc, "'", "''") & "%') "
            SQLQ = SQLQ & "AND TB_NAME = '" & xName & "' "
            rsTemp.Close
            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsTemp.EOF Then
                retVal = rsTemp("TB_KEY")
            End If
            rsTemp.Close
        Else
            If Not rsTemp.EOF Then
                retVal = rsTemp("TB_KEY")
            End If
            rsTemp.Close
        End If
    End If
    getHRTABLCodeFromDesc = retVal
End Function

Public Sub WFCEmployeeDele(xEmpNo As Long)
    Call NukeEE(xEmpNo)
    'Ticket #23116 Franks 01/23/2013 for WFC
    Call NukeEE_SerialNo(xEmpNo)

    'delete HRAUDIT records too
    SQLQ = "DELETE FROM HRAUDIT WHERE AU_EMPNBR = " & xEmpNo
    gdbAdoIhr001X.Execute SQLQ
        
End Sub

Public Function WFCNonUnion(EmpNbr) As Boolean
    Dim rsEmp As New ADODB.Recordset, rsTABL As New ADODB.Recordset
    
    rsEmp.Open "SELECT ED_ORG FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001
    If rsEmp.EOF Then
        WFCNonUnion = False
        rsEmp.Close
        Exit Function
    End If
    If UCase(rsEmp("ED_ORG")) = "NONE" Or UCase(rsEmp("ED_ORG")) = "EXEC" Then WFCNonUnion = True
    rsEmp.Close
End Function

Public Sub WFCTermTransferEmailSending(xEmpNo, xEmpName, xTermDate, xIsTerm As Boolean, Optional xComDIV) 'Ticket #24184 Franks 01/13/2014
    Dim MailBody As String
    Dim LocCode As String, LocDesc As String
    Dim xToEmail As String
    'xIsTerm     -> termination
    'not xIsTerm -> transfer out
    
    On Error GoTo ErrorHandler
    Load frmSendEmail
    'If glbWFC And clpCode(1) = "TOUT" Then   'Ticket #23173 Franks 01/28/2013 - for transfer out
    If Not xIsTerm Then
        frmSendEmail.txtSubject.Text = "info:HR Transfer Notice - " & xEmpName
    Else
        'frmSendEmail.txtSubject.Text = "info:HR Termination Notice"
        'Ticket #18578
        frmSendEmail.txtSubject.Text = "info:HR Termination Notice - " & xEmpName
    End If

    MailBody = "The employee below has been terminated." & vbCrLf & vbCrLf
    MailBody = MailBody & "Employee #: " & xEmpNo & vbCrLf
    MailBody = MailBody & "Name: " & xEmpName & vbCrLf

    If glbWFC Then
        GetLocation xEmpNo, LocCode, LocDesc
        MailBody = MailBody & "Location: " & LocCode & " - " & LocDesc & vbCrLf
        MailBody = MailBody & "Reporting Authority: " & GetReportingAuthority(xEmpNo) & vbCrLf
        'If clpCode(1) = "TOUT" Then
        If Not xIsTerm Then ' Transfer Out
            MailBody = MailBody & "Reason: TOUT - Transfer Out Of Unit" & vbCrLf
            'If comDIV.Visible Then
            If Not IsMissing(xComDIV) Then
                MailBody = MailBody & "Transfer To Division: " & xComDIV & vbCrLf
            End If
        End If
    End If
    MailBody = MailBody & "Date: " & xTermDate & vbCrLf & vbCrLf

    frmSendEmail.txtBody.Text = MailBody
    
    ' dkostka - 02/23/2001 - Automated email sending for WFC.
    If glbWFC Then
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(0).Caption = "Sending email..."
        'Franks 05/03/04 Ticket #6105 David Hili wants to change it
        'frmSendEmail.txtTo.Text = "hotline@woodbridgegroup.com"
        frmSendEmail.txtTo.Text = glbWFCTermEmail '"termnotice@woodbridgegroup.com"
        frmSendEmail.Tag = ""
        frmSendEmail.cmdSend_Click
        Do
            DoEvents
        Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
        ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
        '   otherwise refuse to terminate the employee.
        If frmSendEmail.Tag = "DONE" Then
            Unload frmSendEmail
            AbortTerm = False
        Else
            Unload frmSendEmail
            AbortTerm = True
        End If
        MDIMain.panHelp(0).Caption = ""
        MDIMain.panHelp(0).FloodType = 1
    Else
        frmSendEmail.Show 1
    End If
    
exH:
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH

End Sub

Public Function GetLocation(EmpNbr, ByRef LocCode As String, ByRef LocDesc As String)
    Dim rsEmp As New ADODB.Recordset, rsTABL As New ADODB.Recordset
    
    rsEmp.Open "SELECT ED_LOC FROM HREMP WHERE ED_EMPNBR=" & EmpNbr, gdbAdoIhr001
    If rsEmp.EOF Then
        LocCode = ""
        LocDesc = ""
        rsEmp.Close
        Exit Function
    End If
    If Not IsNull(rsEmp("ED_LOC")) Then
        LocCode = rsEmp("ED_LOC")
    Else
        LocCode = ""
    End If
    rsEmp.Close
    
    rsTABL.Open "SELECT * FROM HRTABL WHERE TB_NAME='EDLC' AND TB_KEY='" & LocCode & "'", gdbAdoIhr001
    If rsTABL.EOF Then
        LocDesc = ""
        rsTABL.Close
        Exit Function
    End If
    LocDesc = rsTABL("TB_DESC")
    rsTABL.Close
End Function

Public Function IsEmailSetup(EmpNbr) As Boolean
    Dim rsEmail As New ADODB.Recordset
    
    rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
    If rsEmail.EOF Then
        IsEmailSetup = False
    Else
        IsEmailSetup = True
    End If
    rsEmail.Close
End Function

Public Function getEmpPosFromReptNo1(xEmpNo)
Dim rsJobHis As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ As String
Dim xRept1Pos, xReptPosCode
Dim retVal
    retVal = ""
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT JH_CURRENT = 0 AND JH_REPTAU = " & xEmpNo & " "
    rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsJobHis.EOF Then
        retVal = rsJobHis("JH_JOB")
    End If
    getEmpPosFromReptNo1 = retVal
End Function

Public Function IsRept1PosNotMatchPosMaster(xRept1EmpNo, xCode)
Dim rsJobHis As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ As String
Dim xRept1Pos, xReptPosCode
Dim retVal
    retVal = False
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xReptPosCode = ""
    If Not rsJOB.EOF Then
        If Not IsNull(rsJOB("JB_REPTAU")) Then
            If Len(rsJOB("JB_REPTAU")) > 0 Then
                xReptPosCode = rsJOB("JB_REPTAU")
            End If
        End If
    End If
    rsJOB.Close
    
    xRept1Pos = ""
    SQLQ = "SELECT JH_EMPNBR,JH_JOB FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & xRept1EmpNo
    rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsJobHis.EOF Then
        xRept1Pos = rsJobHis("JH_JOB")
    End If
    If Not xReptPosCode = xRept1Pos Then
        retVal = True
    End If
    IsRept1PosNotMatchPosMaster = retVal
End Function

Public Function IsRept1PosNotMatchEmpRept1(xEmpNo, xReptNo)
Dim rsJobHis As New ADODB.Recordset
Dim xPosCode
Dim SQLQ As String
Dim retVal As Boolean

retVal = True

'xPosCode = getEmpPosFromReptNo1(xEmpNo) 'Ticket #29484 Franks 11/29/2016
xPosCode = ""
SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT JH_CURRENT = 0 AND JH_EMPNBR = " & xEmpNo & " "
rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsJobHis.EOF Then
    xPosCode = rsJobHis("JH_JOB")
End If
    
If Len(xPosCode) > 0 Then
    If IsRept1PosNotMatchPosMaster(xReptNo, xPosCode) Then
        glbMsgCustomVal = 11
        frmMsgDialog.Show 1
        'if glbMsgCustomVal = 1 then 'If <<Continue>> is checked, save the record with the incorrect RA#1.
        If glbMsgCustomVal = 2 Then 'If <<Cancel>> is checked, undo the change.
            'elpRept(0).Text = GetReportingAuth1EmpNoBasePosMaster(xPosCode)
            retVal = False
        End If
    End If
End If

IsRept1PosNotMatchEmpRept1 = retVal

End Function

Public Function GetReportingAuth1EmpNoBasePosMaster(xCode)
Dim rsEmp As New ADODB.Recordset, rsJobHis As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim SQLQ As String
Dim xReptPosCode
Dim retVal

    retVal = ""
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xCode & "' "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xReptPosCode = ""
    If Not rsJOB.EOF Then
        If Not IsNull(rsJOB("JB_REPTAU")) Then
            If Len(rsJOB("JB_REPTAU")) > 0 Then
                xReptPosCode = rsJOB("JB_REPTAU")
                
                SQLQ = "SELECT TOP 1 JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB = '" & xReptPosCode & "' "
                rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsJobHis.EOF Then
                    If Not IsNull(rsJobHis("JH_REPTAU")) Then
                        retVal = rsJobHis("JH_EMPNBR")
                    End If
                End If
                rsJobHis.Close
                
            End If
        End If
    End If

    GetReportingAuth1EmpNoBasePosMaster = retVal
End Function

Public Function GetReportingAuthority(EmpNbr)
    Dim rsEmp As New ADODB.Recordset, rsJobHis As New ADODB.Recordset
    GetReportingAuthority = ""
    rsJobHis.Open "SELECT JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_EMPNBR=" & EmpNbr, gdbAdoIhr001
    If Not rsJobHis.EOF Then
        If Not IsNull(rsJobHis("JH_REPTAU")) Then
            If IsNumeric(rsJobHis("JH_REPTAU")) Then
                rsEmp.Open "SELECT ED_SURNAME,ED_FNAME FROM HREMP WHERE ED_EMPNBR=" & rsJobHis("JH_REPTAU"), gdbAdoIhr001
                If Not rsEmp.EOF Then
                    GetReportingAuthority = rsEmp("ED_SURNAME") & ", " & rsEmp("ED_FNAME")
                End If
                rsEmp.Close
            End If
        End If
    End If
    rsJobHis.Close
End Function

Public Sub WFC_UptPenDate4WithDOT(xEmpNo, xDOT) 'Ticket #23948 Frank 06/24/2013
'Termination -
'"   Always update Pension Date 4 with the Date of Termination. Remove the check for Eligible for Pension.
Dim rsOther As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HREMP_OTHER WHERE ER_EMPNBR = " & xEmpNo
    If rsOther.State <> 0 Then rsOther.Close
    rsOther.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsOther.EOF Then
        rsOther("ER_PENSIONDATE4") = xDOT
        rsOther.Update
    End If
    rsOther.Close
End Sub

Public Sub WFCUptData2fromDOT(xEmpNo, xdlpDOther2)
Dim SQLQ As String
Dim xID As Long
'If glbWFC And dlpDOther2.Visible Then
    'If IsDate(dlpDOther2.Text) Then  '(dlpTermDate.Text) Then
    If IsDate(xdlpDOther2) Then
        'Ticket #24167 - Getting an error 91. When there is no xNGSStart (assigned in WFCOther2Screen function) the
        'Data2 is not set. So when in this fuction is called the 'Not Data2.Recordset.EOF' gives an error.
        'So I am checking if Data2.RecordSource = "" to avoid the error. I am not too sure of this logic so I am
        'just adding this to avoid error.
        'If Data2.RecordSource <> "" Then
            'If Not Data2.Recordset.EOF Then
                SQLQ = "UPDATE HRBENGRPLIST SET BM_ENDDATE = " & Date_SQL(xdlpDOther2) & " WHERE BM_WRKEMP = '" & glbUserID & "' "
                SQLQ = SQLQ & "AND BM_PCC = 1 "
                gdbAdoIhr001.Execute SQLQ
            'End If
        'End If
    End If
'End If
End Sub

Public Sub WFCUpdateBenefitGroup(xEmpNo) 'Ticket #23247 Franks 07/22/2013
Dim rsBGMST As New ADODB.Recordset
Dim rsBGTMP As New ADODB.Recordset
Dim rsBGEE As New ADODB.Recordset
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String
Dim BelongOldGroup As Boolean
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001W.CommitTrans

    gdbAdoIhr001W.BeginTrans
    SQLQ = "SELECT * FROM HRBENGRPLIST WHERE BM_WRKEMP = '" & glbUserID & "' "
    rsBGTMP.Open SQLQ, gdbAdoIhr001W, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
    'SQLQ = SQLQ & "AND BF_PCC = 1 " 'Paid Benefits only
    SQLQ = SQLQ & "ORDER BY BF_BCODE, BF_EDATE "

    rsBGMST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    Do While Not rsBGMST.EOF
        rsBGTMP.AddNew
        rsBGTMP("BM_COMPNO") = "001"
        rsBGTMP("BM_BENEFIT_GROUP") = rsBGMST("BF_GROUP")
        rsBGTMP("BM_BCODE") = rsBGMST("BF_BCODE")
        rsBGTMP("BM_EDATE") = rsBGMST("BF_EDATE")
        rsBGTMP("BM_ENDDATE") = rsBGMST("BF_CEASEDATE") 'New
        rsBGTMP("BM_CHECK") = 1
        rsBGTMP("BM_COVER") = rsBGMST("BF_COVER")
        rsBGTMP("BM_AMT") = rsBGMST("BF_AMT")
        rsBGTMP("BM_PPAMT") = rsBGMST("BF_PPAMT")
        rsBGTMP("BM_UNITCOST") = rsBGMST("BF_UNITCOST")
        rsBGTMP("BM_PCE") = rsBGMST("BF_PCE")
        rsBGTMP("BM_PCC") = rsBGMST("BF_PCC")
        rsBGTMP("BM_ECOST") = rsBGMST("BF_ECOST")
        rsBGTMP("BM_CCOST") = rsBGMST("BF_CCOST")
        rsBGTMP("BM_TCOST") = rsBGMST("BF_TCOST")
        rsBGTMP("BM_MAXDOL") = rsBGMST("BF_MAXDOL")
        rsBGTMP("BM_PREMIUM") = rsBGMST("BF_PREMIUM")
        rsBGTMP("BM_PER") = rsBGMST("BF_PER")
        rsBGTMP("BM_MTHCCOST") = rsBGMST("BF_MTHCCOST")
        rsBGTMP("BM_MTHECOST") = rsBGMST("BF_MTHECOST")
        rsBGTMP("BM_TAXBEN") = rsBGMST("BF_TAXBEN")
        rsBGTMP("BM_SALARYDEPENDANT") = rsBGMST("BF_SALARYDEPENDANT")
        rsBGTMP("BM_MINIMUM") = rsBGMST("BF_MINIMUM")
        rsBGTMP("BM_FACTOR") = rsBGMST("BF_FACTOR")
        rsBGTMP("BM_ROUND") = rsBGMST("BF_ROUND")
        rsBGTMP("BM_MAXIMUM") = rsBGMST("BF_MAXIMUM")
        rsBGTMP("BM_NEXTNEAREST") = rsBGMST("BF_NEXTNEAREST")
        rsBGTMP("BM_TAXAMOUNT") = rsBGMST("BF_TAXAMOUNT")
        rsBGTMP("BM_WAITPERIOD") = rsBGMST("BF_WAITPERIOD")
        rsBGTMP("BM_DWM") = rsBGMST("BF_DWM")
        rsBGTMP("BM_PERORDOLL") = rsBGMST("BF_PERORDOLL")
        rsBGTMP("BM_POLICY") = rsBGMST("BF_POLICY")
        rsBGTMP("BM_RATELEVEL") = rsBGMST("BF_RATELEVEL")
        rsBGTMP("BM_COMMENTS") = rsBGMST("BF_COMMENTS")
        rsBGTMP("BM_PTAX") = rsBGMST("BF_PTAX")
        rsBGTMP("BM_ACTION") = "Add"
        rsBGTMP("BM_WRKEMP") = glbUserID
        
        SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'BNCD' AND TB_KEY = '" & rsBGMST("BF_BCODE") & "' "
        rsTABL.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
        If Not rsTABL.EOF Then
            rsBGTMP("BM_BCODE_DESC") = rsTABL("TB_DESC")
        End If
        rsTABL.Close
        rsBGTMP.Update
        rsBGMST.MoveNext
    Loop
    rsBGTMP.Close
    rsBGMST.Close
    gdbAdoIhr001W.CommitTrans
    Call Pause(1)

End Sub

Private Sub WFC_NGSBenEndDateUpPub(xEmpNo, xLastDate) 'Ticket #23247 Franks 07/22/2013
Dim SQLQ, xACT
Dim rsBN As New ADODB.Recordset
Dim rsEmpBN As New ADODB.Recordset
Dim xTemp
Dim xDate1, xDate2
    'update Last Day
    If IsDate(xLastDate) Then
        SQLQ = "UPDATE HREMP SET ED_LDAY = " & Date_SQL(xLastDate) & " "
        SQLQ = SQLQ & " WHERE ED_EMPNBR = " & xEmpNo 'Ticket #24588 Franks 11/01/2013
        gdbAdoIhr001.Execute SQLQ
        Call WFCAUDITBENF_NGSEnd(xEmpNo, False, , "Y", xLastDate)
    End If
    
    SQLQ = "SELECT * FROM HRBENGRPLIST "
    SQLQ = SQLQ & "WHERE BM_WRKEMP = '" & glbUserID & "'  "
    rsBN.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsBN.EOF
        SQLQ = "SELECT * FROM HRBENFT WHERE  BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_BCODE = '" & rsBN("BM_BCODE") & "' "
        If Not IsNull(rsBN("BM_EDATE")) Then SQLQ = SQLQ & "AND BF_EDATE = " & Date_SQL(rsBN("BM_EDATE")) & " "
        If rsEmpBN.State <> 0 Then rsEmpBN.Close
        rsEmpBN.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsEmpBN.EOF Then
            If IsNull(rsEmpBN("BF_CEASEDATE")) Then xDate1 = CVDate("01/01/1900") Else xDate1 = CVDate(rsEmpBN("BF_CEASEDATE"))
            If IsNull(rsBN("BM_ENDDATE")) Then xDate2 = CVDate("01/01/1900") Else xDate2 = CVDate(rsBN("BM_ENDDATE"))
            rsEmpBN("BF_CEASEDATE") = rsBN("BM_ENDDATE")
            rsEmpBN.Update
            If Not xDate1 = xDate2 Then 'BF_CEASEDATE was changed
                If xDate2 > CVDate("01/01/1900") Then
                    'update hraudit - begin
                    Call WFCAUDITBENF_NGSEnd(xEmpNo, False, rsEmpBN)
                    'update hraudit - end
                End If
            End If
        End If
        rsEmpBN.Close
        rsBN.MoveNext
    Loop
    rsBN.Close
End Sub

'09/02/2015 non WFC Diff Benefit End Date
Public Function NonWFCAUDITBENF_End(xEmpNo, xlocNewRec As Boolean, Optional rslBen As ADODB.Recordset, Optional xIsWorkDay = "N", Optional xLastDate)
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
Dim ACTX
Dim NBCode, NPPAMT, NMTHCOMP, NMTHEMP, NBAMT, NPPE, NPCC, NMAXDOL, NEDate, NCOVER, NTCOST
Dim xTermSEQ
Dim SQLQ As String

'''On Error GoTo AUDIT_ERR
NonWFCAUDITBEN_End = False

If xlocNewRec Then
    ACTX = "A"
Else
    ACTX = "M"
End If

xTermSEQ = 0
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
Else
    SQLQ = "SELECT ED_PT,ED_DIV FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDIV = ""
    Else
        xDIV = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDIV = ""
End If
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS,AU_CEASEDATE,AU_LDAY "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

xADD = False

If xIsWorkDay = "N" Then
    NBCode = ""
    NPPAMT = ""
    NMTHCOMP = ""
    NMTHEMP = ""
    NBAMT = ""
    NPPE = ""
    NPCC = ""
    NMAXDOL = ""
    NEDate = ""
    NCOVER = ""
    NTCOST = ""
    NBCode = rslBen("BF_BCODE")
    If Not IsNull(rslBen("BF_EDATE")) Then NEDate = rslBen("BF_EDATE")
End If

'GoTo MODNOUPD

'BF_CEASEDATE was changed
MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDIV

If xIsWorkDay = "N" Then
    rsTA("AU_BCODE") = NBCode 'clpCode(1).Text
    rsTA("AU_CEASEDATE") = rslBen("BF_CEASEDATE")
    rsTA("AU_LDATE") = Date
    If IsDate(NEDate) Then 'if benefit effe date is future date, use it as LDATE
        If CVDate(NEDate) > CVDate(Date) Then
            rsTA("AU_LDATE") = CVDate(NEDate)
        End If
    End If
End If
If xIsWorkDay = "Y" Then
    rsTA("AU_LDAY") = xLastDate
    rsTA("AU_LDATE") = Date
End If
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
Else
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update
rsTA.Close

MODNOUPD:
NonWFCAUDITBEN_End = True
Exit Function
AUDIT_ERR:


End Function

Public Function WFCAUDITBENF_NGSEnd(xEmpNo, xlocNewRec As Boolean, Optional rslBen As ADODB.Recordset, Optional xIsWorkDay = "N", Optional xLastDate)
Dim rsEmp As New ADODB.Recordset
Dim rsTA As New ADODB.Recordset
Dim rsTB As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim strFields As String
Dim ACTX
Dim NBCode, NPPAMT, NMTHCOMP, NMTHEMP, NBAMT, NPPE, NPCC, NMAXDOL, NEDate, NCOVER, NTCOST
Dim xTermSEQ
Dim SQLQ As String

'''On Error GoTo AUDIT_ERR
WFCAUDITBENF_NGSEnd = False

If xlocNewRec Then
    ACTX = "A"
Else
    ACTX = "M"
End If

xTermSEQ = 0
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
Else
    SQLQ = "SELECT ED_PT,ED_DIV FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsTB.Open SQLQ, gdbAdoIhr001, adOpenKeyset

If Not rsTB.EOF Then
    If IsNull(rsTB("ED_PT")) Then
        xPT = ""
    Else
        xPT = rsTB("ED_PT")
    End If
    If IsNull(rsTB("ED_DIV")) Then
        xDIV = ""
    Else
        xDIV = rsTB("ED_DIV")
    End If
Else
    xPT = ""
    xDIV = ""
End If
'strfields added by Bryan 02/Dec/05 Ticket#9899
strFields = "AU_LOC_TABL, AU_SECTION_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL, "
strFields = strFields & "AU_DOLENT_TABL, AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_COVER, AU_EDATE, AU_MAXDOL, AU_PPAMT, "
strFields = strFields & "AU_MTHCCOST, AU_MTHECOST, AU_BCODE, AU_BNAME, AU_BRELATE, AU_BDOB, AU_TAXBEN, AU_COVER, AU_TCOST, AU_PREMIUM, AU_PCE, AU_PCC, "
strFields = strFields & "AU_OLDPPMT, AU_MAXDOL, AU_EDATE, AU_PER, AU_BAMT, AU_UNITCOST, AU_BCODE, AU_BNAME, "
strFields = strFields & "AU_BRELATE, AU_BDOB, AU_PAYROLL_ID, AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE,AU_OLDLOC,AU_OLDWHRS,AU_CEASEDATE,AU_LDAY "
rsTA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001, adOpenKeyset, adLockOptimistic

xADD = False

If xIsWorkDay = "N" Then
    NBCode = ""
    NPPAMT = ""
    NMTHCOMP = ""
    NMTHEMP = ""
    NBAMT = ""
    NPPE = ""
    NPCC = ""
    NMAXDOL = ""
    NEDate = ""
    NCOVER = ""
    NTCOST = ""
    NBCode = rslBen("BF_BCODE")
    If Not IsNull(rslBen("BF_EDATE")) Then NEDate = rslBen("BF_EDATE")
    ''If Not IsNull(rslBen("BF_PPAMT")) Then NPPAMT = rslBen("BF_PPAMT")
    ''If Not IsNull(rslBen("BF_MTHCCOST")) Then NMTHCOMP = rslBen("BF_MTHCCOST")
    ''If Not IsNull(rslBen("BF_MTHECOST")) Then NMTHEMP = rslBen("BF_MTHECOST")
    ''If Not IsNull(rslBen("BF_AMT")) Then NBAMT = rslBen("BF_AMT")
    ''If Not IsNull(rslBen("BF_PCC")) Then NPCC = rslBen("BF_PCC")
    ''If Not IsNull(rslBen("BF_PCE")) Then NPPE = rslBen("BF_PCE")
    ''If Not IsNull(rslBen("BF_MAXDOL")) Then NMAXDOL = rslBen("BF_MAXDOL")
    ''If Not IsNull(rslBen("BF_COVER")) Then NCOVER = rslBen("BF_COVER")
    ''If Not IsNull(rslBen("BF_TCOST")) Then NTCOST = rslBen("BF_TCOST")
    ''
    ''If OBCode <> NBCode Then GoTo MODUPD
    '''If OPPE <> NPPE Or OPCC <> NPCC Then GoTo MODUPD
    ''If OPPAMT <> NPPAMT Or OMAXDOL <> NMAXDOL Then GoTo MODUPD
    '''If OMTHCOMP <> NMTHCOMP Or OMTHEMP <> NMTHEMP Then GoTo MODUPD
    ''If OBAMT <> NBAMT Then GoTo MODUPD
    ''If OEDate <> NEDate Then GoTo MODUPD
End If

'GoTo MODNOUPD

'BF_CEASEDATE was changed
MODUPD:

rsTA.AddNew
rsTA("AU_LOC_TABL") = "EDLC": rsTA("AU_SECTION_TABL") = "EDSE": rsTA("AU_EMP_TABL") = "EDEM": rsTA("AU_SUPCODE_TABL") = "EDSP"
rsTA("AU_ORG_TABL") = "EDOR": rsTA("AU_PAYP_TABL") = "SDPP": rsTA("AU_BCODE_TABL") = "BNCD": rsTA("AU_TREAS_TABL") = "TERM"
rsTA("AU_DOLENT_TABL") = "EDOL": rsTA("AU_EARN_TABL") = "EARN"
rsTA("AU_NEWEMP") = "N"
rsTA("AU_PTUPL") = xPT
rsTA("AU_DIVUPL") = xDIV

If xIsWorkDay = "N" Then
    rsTA("AU_BCODE") = NBCode 'clpCode(1).Text
    rsTA("AU_CEASEDATE") = rslBen("BF_CEASEDATE")
    'If OMTHCOMP <> NMTHCOMP Then rsTA("AU_MTHCCOST") = NMTHCOMP
    'If OMTHEMP <> NMTHEMP Then rsTA("AU_MTHECOST") = NMTHEMP
    'If OTAXBEN <> txtTAXBEN Then rsTA("AU_TAXBEN") = txtTAXBEN
    'If OCOVER <> NCOVER Then rsTA("AU_COVER") = NCOVER
    'If OTCOST <> NTCOST Then rsTA("AU_TCOST") = NTCOST
    'If OPremium <> lblAP Then rsTA("AU_PREMIUM") = lblAP
    'If OPPE <> NPPE Then rsTA("AU_PCE") = NPPE
    'If OPCC <> NPCC Then rsTA("AU_PCC") = NPCC
    'If OPPAMT <> NPPAMT Then
    '    rsTA("AU_PPAMT") = NPPAMT
    '    If IsNumeric(OPPAMT) Then rsTA("AU_OLDPPMT") = Val(OPPAMT)
    'End If
    'If OMAXDOL <> NMAXDOL Then rsTA("AU_MAXDOL") = NMAXDOL
    'If OEDate <> NEDate Then
    '  If IsDate(NEDate) Then
    '      rsTA("AU_EDATE") = CVDate(NEDate)
    '  End If
    'End If
    'If OPER <> txtPer Then rsTA("AU_PER") = txtPer
    'If OBAMT <> NBAMT Then rsTA("AU_BAMT") = NBAMT
    'If OUNITCOST <> medUnitCost Then rsTA("AU_UNITCOST") = IIf(medUnitCost = "", 0, medUnitCost)
    rsTA("AU_LDATE") = Date
    If IsDate(NEDate) Then 'if benefit effe date is future date, use it as LDATE
        If CVDate(NEDate) > CVDate(Date) Then
            rsTA("AU_LDATE") = CVDate(NEDate)
        End If
    End If
End If
If xIsWorkDay = "Y" Then
    rsTA("AU_LDAY") = xLastDate
    rsTA("AU_LDATE") = Date
End If
If xTermSEQ = 0 Then
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
Else
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND TERM_SEQ = " & xTermSEQ & " "
End If
rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsEmp.EOF Then
    If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then rsTA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
End If
rsEmp.Close
rsTA("AU_COMPNO") = "001"
rsTA("AU_EMPNBR") = xEmpNo
rsTA("AU_LUSER") = glbUserID
rsTA("AU_LTIME") = Time$
rsTA("AU_UPLOAD") = "N"
rsTA("AU_TYPE") = ACTX
rsTA.Update
rsTA.Close

MODNOUPD:
WFCAUDITBENF_NGSEnd = True
Exit Function
AUDIT_ERR:

End Function

Public Function Get_WFC_COMPA_FromMaster(xUnion, xJob, xSalary, fglbSection, xMarketline, xFiscalYear) 'Ticket #25045 Franks 02/05/2014
Dim xDollear
Dim lblsalstate0, lblsalstate1, lblsalstate2
Dim rsJOB As New ADODB.Recordset
Dim rsWFC As New ADODB.Recordset
Dim fglbBAND ', fglbSection, xFiscalYear, xMarketLine
Dim X%, I%
Dim xItemAdd
Dim SQLQ
Dim retVal
    
retVal = 0
If xUnion = "NONE" Or xUnion = "EXEC" Then
    lblsalstate0 = 0
    lblsalstate1 = 0
    lblsalstate2 = 0
    
    If IsNull(fglbSection) Then fglbSection = ""
    If IsNull(xMarketline) Then xMarketline = ""
    If IsNull(xFiscalYear) Then xFiscalYear = 0
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xJob & "' "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsJOB.EOF Then
        If Not IsNull(rsJOB("JB_BAND")) Then fglbBAND = rsJOB("JB_BAND") Else fglbBAND = "*"

        SQLQ = "select * from WFC_Salary_Administration "
        SQLQ = SQLQ & " WHERE [BAND]='" & fglbBAND & "'"
        If Len(fglbSection) > 0 Then
            SQLQ = SQLQ & " AND SectionCode ='" & fglbSection & "'"
        End If
        If Len(xFiscalYear) > 0 Then
            SQLQ = SQLQ & " AND FiscalYear =" & xFiscalYear & ""
        End If
        'SQLQ = SQLQ & " group by MarketLine"
        SQLQ = SQLQ & " ORDER by MarketLine"
        rsWFC.Open SQLQ, gdbAdoIhrWFC, adOpenKeyset
        X% = 0
        'cmbMarketLine.Clear
        Do Until rsWFC.EOF
            'cmbMarketLine.AddItem rsWFC("marketline")
            If rsWFC("marketline") = xMarketline Then
                'cmbMarketLine.ListIndex = X%
                lblsalstate0 = rsWFC("LDOLLARS")
                lblsalstate1 = rsWFC("MDOLLARS")
                lblsalstate2 = rsWFC("HDOLLARS")
            End If
            X% = X% + 1
            rsWFC.MoveNext
        Loop
        rsWFC.Close

        If IsNumeric(lblsalstate1) Then xDollear = lblsalstate1 Else xDollear = 0
        If Val(xDollear) <> 0 Then
            If IsNumeric(xSalary) And xDollear > 0 Then
                retVal = (Val(xSalary) / xDollear) * 100
            End If
        End If
        If retVal > 999.99 Then retVal = "999.99"
    End If
End If

Get_WFC_COMPA_FromMaster = retVal
    
End Function

Public Function getGLDesc(xCode)
Dim rsGL As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT * FROM HRGL WHERE GL_NO = '" & xCode & "' "
        rsGL.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsGL.EOF Then
            xRetVal = rsGL("GL_DESCR")
        End If
        rsGL.Close
    End If
    getGLDesc = xRetVal
End Function

Public Function getNewJobCodeDescPub(xKey)
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String, xStr As String
    SQLQ = "SELECT * FROM HRJOBMASTER WHERE JB_JOBCODE = '" & xKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    xStr = ""
    If Not rsTABL.EOF Then
        xStr = rsTABL("JB_JOBDESCR")
    End If
    rsTABL.Close
    getNewJobCodeDescPub = Left(xStr, 100)
End Function

Public Function getNewJobMasterCode(xKey) 'Ticket #27531 Franks 09/21/2015
Dim rsTABL As New ADODB.Recordset
Dim SQLQ As String, xStr As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT JB_CODE,JB_JOBCODE FROM HRJOB WHERE JB_CODE = '" & xKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTABL.EOF Then
        retVal = rsTABL("JB_JOBCODE")
    End If
    rsTABL.Close
    getNewJobMasterCode = retVal
End Function

Public Function getDivDescPub(xCode)
Dim rsDiv As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT DIV,Division_Name FROM HR_DIVISION WHERE DIV = '" & xCode & "' "
        rsDiv.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDiv.EOF Then
            xRetVal = rsDiv("Division_Name")
        End If
        rsDiv.Close
    End If
    getDivDescPub = xRetVal
End Function

Public Function getDeptDescPub(xCode)
Dim rsDEPT As New ADODB.Recordset
Dim SQLQ, xRetVal
    xRetVal = ""
    If Not IsNull(xCode) Then
        SQLQ = "SELECT DF_NBR, DF_NAME FROM HRDEPT WHERE DF_NBR = '" & xCode & "' "
        rsDEPT.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsDEPT.EOF Then
            xRetVal = rsDEPT("DF_NAME")
        End If
        rsDEPT.Close
    End If
    getDeptDescPub = xRetVal
End Function

Public Function GetTABLCodePub(xName, xCode) 'Ticket #16544
Dim rsTABL As New ADODB.Recordset
Dim xStr As String
    xRetVal = ""
    If Not IsNull(xCode) Then
        rsTABL.Open "SELECT * FROM HRTABL WHERE TB_NAME='" & xName & "' AND TB_KEY='" & xCode & "'", gdbAdoIhr001, adOpenStatic, adLockPessimistic
        If Not rsTABL.EOF Then
            xRetVal = rsTABL("TB_DESC")
        End If
        rsTABL.Close
    End If
GetTABLCodePub = xRetVal
End Function

Public Sub WFCUpdBenefitEndDatePublic(xEmpNo, xDATE, xType) 'Ticket #24179 Franks 02/25/2014
Dim rsBenT As New ADODB.Recordset
Dim SQLQ As String
    If Not IsDate(xDATE) Then
        Exit Sub
    End If
    If xType = "ALL" Then
        If glbMsgCustomVal = 4 Then 'remove Benefit Group
            'clpBGroup.Text = ""
            SQLQ = "UPDATE HREMP SET ED_BENEFIT_GROUP = NULL WHERE ED_EMPNBR = " & xEmpNo & " "
            gdbAdoIhr001.Execute SQLQ
        End If
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_PCC = 1 "
        SQLQ = SQLQ & "AND (BF_CEASEDATE IS NULL) "
        If rsBenT.State <> 0 Then rsBenT.Close
        rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsBenT.EOF
            rsBenT("BF_CEASEDATE") = CVDate(xDATE)
            rsBenT.Update
            'update audit
            Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
            rsBenT.MoveNext
        Loop
        rsBenT.Close
    End If
    If xType = "ComPaidNoIE" Then
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_PCC = 1 "
        SQLQ = SQLQ & "AND NOT (BF_BCODE = 'IE') "
        SQLQ = SQLQ & "AND (BF_CEASEDATE IS NULL) "
        If rsBenT.State <> 0 Then rsBenT.Close
        rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsBenT.EOF
            rsBenT("BF_CEASEDATE") = CVDate(xDATE)
            rsBenT.Update
            'update audit
            Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
            rsBenT.MoveNext
        Loop
        rsBenT.Close
    End If
    If xType = "RemoveEndDate" Then
        SQLQ = "SELECT * FROM HRBENFT WHERE BF_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND BF_PCC = 1 "
        SQLQ = SQLQ & "AND not (BF_CEASEDATE IS NULL) "
        If rsBenT.State <> 0 Then rsBenT.Close
        rsBenT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        Do While Not rsBenT.EOF
            rsBenT("BF_CEASEDATE") = Null ' CVDate(xDate)
            rsBenT.Update
            'update audit
            Call WFC_AUDITBEN_ByField(xEmpNo, "M", "BF_CEASEDATE", rsBenT)
            rsBenT.MoveNext
        Loop
        rsBenT.Close
    End If
End Sub

Public Function getWFCRA4(xDIV)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    If Len(xDIV) > 0 Then
        SQLQ = "SELECT * FROM HR_DIVISION WHERE DIV = '" & xDIV & "' "
        If rs.State <> 0 Then rs.Close
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rs.EOF Then
            If Not IsNull(rs("DV_BONUSDEPT")) Then
                If IsNumeric(rs("DV_BONUSDEPT")) Then
                    retVal = rs("DV_BONUSDEPT")
                End If
            End If
        End If
        rs.Close
    End If
    getWFCRA4 = retVal
End Function

Public Function WFCHRSoftMissNewhire(xEmpNo, xUptField)
Dim rsLEmp As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = False
    'find candidate
    SQLQ = "SELECT ED_EMPNBR, ED_CANDIDATE FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    If rsLEmp.State <> 0 Then rsLEmp.Close
    rsLEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsLEmp.EOF Then
        If Not IsNull(rsLEmp("ED_CANDIDATE")) Then
            If rsLEmp("ED_CANDIDATE") > 0 Then
                'SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_HIRETYPE = 'NEW' AND SF_EMPNBR = " & xEmpNo & " "
                SQLQ = "SELECT * FROM HRSF_XML_IMPORT WHERE SF_HIRETYPE = 'NEW' AND SF_CANDIDATE = " & rsLEmp("ED_CANDIDATE") & " "
                If rs.State <> 0 Then rs.Close
                rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rs.EOF Then
                    If rs(xUptField) = False Then
                        retVal = True
                        glbCandidate = rsLEmp("ED_CANDIDATE")
                    End If
                End If
                rs.Close
            End If
        End If
    End If
    rsLEmp.Close
    WFCHRSoftMissNewhire = retVal
End Function

Public Function getWFCPosSec(xAPos)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT JB_CODE, JB_SECTION FROM HRJOB WHERE JB_CODE ='" & xAPos & "' "
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        If Not IsNull(rs("JB_SECTION")) Then
            retVal = rs("JB_SECTION")
        End If
    End If
    rs.Close
    getWFCPosSec = retVal
End Function

Public Function mod_ADD_Pos_Budget_WFC(xAPos, xASection, xDIV, xBUDGNBR, xUserID) 'Ticket #29484 Franks 11/21/2016
Dim snapJobCount As New ADODB.Recordset
Dim rsHRJOB As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&, spct%
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#, FTEHrs#
Dim snapFTENum As New ADODB.Recordset
Dim snapFTEHrs As New ADODB.Recordset
Dim snapBudget As New ADODB.Recordset
Dim xJob, xDeptno, xGLNO, xPosCtrl
Dim xSec
Dim xYear

On Error GoTo mod_Upd_Pos_Budget_WFC_Err
pct# = 1

SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
SQLQ = SQLQ & "AND JG_CURRENT = 1 "
If Not Len(xAPos) = 0 Then SQLQ = SQLQ & "AND JG_CODE = '" & xAPos & "' "
'If Not Len(xASection) = 0 Then SQLQ = SQLQ & "AND JG_SECTION = '" & xASection & "' "

If snapBudget.State <> 0 Then snapBudget.Close
snapBudget.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic

If Not snapBudget.EOF Then
    Exit Function 'found it then exit, this function is adding new only
End If

snapBudget.AddNew
snapBudget("JG_COMPNO") = "001"
snapBudget("JG_CODE") = xAPos
snapBudget("JG_SECTION") = xASection
snapBudget("JG_DIV") = xDIV
'snapBudget("JG_BUDPOSNBR") = "" '?
snapBudget("JG_BUDGNBR") = xBUDGNBR
snapBudget("JG_NBRFIL") = 0
snapBudget("JG_VACANCY_POS") = xBUDGNBR - 0
snapBudget("JG_FTENUM") = xBUDGNBR
snapBudget("JG_FTENUMFILL") = 0
snapBudget("JG_FTENUMVACN") = xBUDGNBR - 0
If Mid(xAPos, 5, 3) = "IND" Then 'Ticket #30358 Franks 07/13/2017 - leave these as blank for Independent Contractor Positions
    snapBudget("JG_FTEHRS") = 0
    snapBudget("JG_FTETOTHR") = 0
Else
    snapBudget("JG_FTEHRS") = 2080
    snapBudget("JG_FTETOTHR") = 2080
End If
snapBudget("JG_LDATE") = Date
snapBudget("JG_LTIME") = Time$
snapBudget("JG_LUSER") = xUserID

If month(Date) = 11 Or month(Date) = 12 Then
    xYear = Year(Date) + 1
Else
    xYear = Year(Date)
End If
snapBudget("JG_YEAR") = xYear
SQLQ = "SELECT * FROM HRPARCO"
If rs.State <> 0 Then rs.Close
rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rs.EOF Then
    If Not IsNull(rs("PC_TDATE")) Then
        I = xYear - Year(rs("PC_TDATE"))
        snapBudget("JG_FRDATE") = DateAdd("YYYY", I, rs("PC_FDATE"))
        snapBudget("JG_TODATE") = DateAdd("YYYY", I, rs("PC_TDATE"))
        snapBudget("JG_EFDATE") = snapBudget("JG_FRDATE")
    End If
End If
rs.Close
snapBudget("JG_CURRENT") = 1
snapBudget.Update
snapBudget.Close

Exit Function


mod_Upd_Pos_Budget_WFC_Err:
If Err = 94 Then
Err = 0
Resume Next
End If
glbFrmCaption$ = "Module - Add Budgeted Positions"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update HRJOBBUD Count", "HRJOBBUD", "Add")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function


Public Function mod_Upd_Pos_Budget_WFC(xAPos, xASection, Optional xEmpNo)
Dim snapJobCount As New ADODB.Recordset
Dim rsHRJOB As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&, spct%
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#, FTEHrs#
Dim snapFTENum As New ADODB.Recordset
Dim snapFTEHrs As New ADODB.Recordset
Dim snapBudget As New ADODB.Recordset
Dim xJob, xDIV, xDeptno, xGLNO, xPosCtrl
Dim xSec

mod_Upd_Pos_Budget_WFC = False
On Error GoTo mod_Upd_Pos_Budget_WFC_Err
MDIMain.panHelp(0).FloodShowPct = True
MDIMain.panHelp(0).ForeColor = &HFFFFFF
pct# = 1

MDIMain.panHelp(0).FloodType = 1

If Not IsMissing(xEmpNo) Then
    xAPos = getEmpPostion(xEmpNo)
End If

If Len(xAPos) > 0 Then
    If Len(xASection) = 0 Then
        xASection = getWFCPosSec(xAPos)
        If Len(xASection) = 0 Then
            Exit Function
        End If
    End If
End If

SQLQ = "SELECT * FROM HRJOBBUD WHERE (1=1) "
'If Len(xAPos) = 0 Then
    SQLQ = SQLQ & "AND JG_CURRENT = 1 "
'End If
If Not Len(xAPos) = 0 Then SQLQ = SQLQ & "AND JG_CODE = '" & xAPos & "' "
If Not Len(xASection) = 0 Then SQLQ = SQLQ & "AND JG_SECTION = '" & xASection & "' "

If snapBudget.State <> 0 Then snapBudget.Close
snapBudget.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not (snapBudget.EOF And snapBudget.BOF) Then
    snapBudget.MoveLast
    rcount& = snapBudget.RecordCount
    snapBudget.MoveFirst
End If
pct# = 0
Do While Not snapBudget.EOF
    MDIMain.panHelp(0).FloodPercent = (pct# / rcount&) * 100
    pct# = pct# + 1
    xDIV = "": xDeptno = "": xGLNO = "": xPosCtrl = "": xSec = ""
    xJob = snapBudget("JG_CODE")
    If Not IsNull(snapBudget("JG_SECTION")) Then
        xSec = snapBudget("JG_SECTION")
    End If
    If Not IsNull(snapBudget("JG_DIV")) Then
        xDIV = snapBudget("JG_DIV")
    End If
    If Not IsNull(snapBudget("JG_DEPTNO")) Then
        xDeptno = snapBudget("JG_DEPTNO")
    End If
    If Not IsNull(snapBudget("JG_GLNO")) Then
        xGLNO = snapBudget("JG_GLNO")
    End If
    'Position filled

    'SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB, "
    SQLQ = "SELECT HR_JOB_HISTORY.JH_JOB, "
    If Len(xSec) > 0 Then SQLQ = SQLQ & "HREMP.ED_SECTION, "
    If Len(xDIV) > 0 Then SQLQ = SQLQ & "HREMP.ED_DIV, "
    If Len(xDeptno) > 0 Then SQLQ = SQLQ & "HREMP.ED_DEPTNO, "
    If Len(xGLNO) > 0 Then SQLQ = SQLQ & "HREMP.ED_GLNO, "
    SQLQ = SQLQ & "COUNT(HR_JOB_HISTORY.JH_EMPNBR) AS NoPosFilled  "
    SQLQ = SQLQ & "FROM HR_JOB_HISTORY "
    SQLQ = SQLQ & "INNER JOIN HREMP ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJob & "' "
    SQLQ = SQLQ & "AND NOT HREMP.ED_EMP = 'RET' " 'exclude employees the RET Employment Status
    SQLQ = SQLQ & "AND NOT (ED_FNAME LIKE '%(Deceased)%') " 'not include Death employees
    If Len(xSec) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_SECTION = '" & xSec & "' "
    If Len(xDIV) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_DIV = '" & xDIV & "' "
    If Len(xDeptno) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_DEPTNO = '" & xDeptno & "' "
    If Len(xGLNO) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_GLNO = '" & xGLNO & "' "
    'SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB "
    SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_JOB "
    If Len(xSec) > 0 Then SQLQ = SQLQ & ",HREMP.ED_SECTION "
    If Len(xDIV) > 0 Then SQLQ = SQLQ & ",HREMP.ED_DIV "
    If Len(xDeptno) > 0 Then SQLQ = SQLQ & ",HREMP.ED_DEPTNO "
    If Len(xGLNO) > 0 Then SQLQ = SQLQ & ",HREMP.ED_GLNO "

    snapJobCount.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not snapJobCount.EOF Then
        If Not IsNull(snapJobCount("NoPosFilled")) Then
            JobCount& = snapJobCount("NoPosFilled")
        Else
            JobCount& = 0
        End If
        snapBudget("JG_NBRFIL") = JobCount&
        snapBudget("JG_VACANCY_POS") = snapBudget("JG_BUDGNBR") - JobCount&
        snapBudget.Update
    Else
        snapBudget("JG_NBRFIL") = 0
        snapBudget("JG_VACANCY_POS") = snapBudget("JG_BUDGNBR") '- JobCount&
    End If
    snapJobCount.Close

    'FTE # filled & FTE Hours/Year
    'SQLQ = "SELECT HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB, "
    SQLQ = "SELECT HR_JOB_HISTORY.JH_JOB, "
    If Len(xSec) > 0 Then SQLQ = SQLQ & "HREMP.ED_SECTION, "
    If Len(xDIV) > 0 Then SQLQ = SQLQ & "HREMP.ED_DIV, "
    If Len(xDeptno) > 0 Then SQLQ = SQLQ & "HREMP.ED_DEPTNO, "
    If Len(xGLNO) > 0 Then SQLQ = SQLQ & "HREMP.ED_GLNO, "
    SQLQ = SQLQ & "SUM(JH_FTENUM) AS FTENumTot, SUM(JH_FTEHRS) AS FTEHrsTot "
    SQLQ = SQLQ & "FROM HR_JOB_HISTORY  "
    SQLQ = SQLQ & "INNER JOIN HREMP ON HREMP.ED_EMPNBR = HR_JOB_HISTORY.JH_EMPNBR "
    SQLQ = SQLQ & "WHERE (JH_CURRENT <> 0) AND JH_JOB = '" & xJob & "' "
    If Len(xSec) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_SECTION = '" & xSec & "' "
    If Len(xDIV) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_DIV = '" & xDIV & "' "
    If Len(xDeptno) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_DEPTNO = '" & xDeptno & "' "
    If Len(xGLNO) > 0 Then SQLQ = SQLQ & "AND HREMP.ED_GLNO = '" & xGLNO & "' "
    SQLQ = SQLQ & "AND NOT (ED_EMP = 'RET') " 'not include RET employees
    SQLQ = SQLQ & "AND NOT (ED_FNAME LIKE '%(Deceased)%') " 'not include Death employees
    'SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_COMPNO, HR_JOB_HISTORY.JH_JOB"
    SQLQ = SQLQ & "GROUP BY HR_JOB_HISTORY.JH_JOB"
    If Len(xSec) > 0 Then SQLQ = SQLQ & ",HREMP.ED_SECTION "
    If Len(xDIV) > 0 Then SQLQ = SQLQ & ",HREMP.ED_DIV "
    If Len(xDeptno) > 0 Then SQLQ = SQLQ & ",HREMP.ED_DEPTNO "
    If Len(xGLNO) > 0 Then SQLQ = SQLQ & ",HREMP.ED_GLNO "

    '--------
    snapFTENum.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not snapFTENum.EOF Then
        If Not IsNull(snapFTENum("FTENumTot")) Then
            FTENum# = snapFTENum("FTENumTot")
        Else
            FTENum# = 0
        End If
        If Not IsNull(snapFTENum("FTEHrsTot")) Then
            FTEHrs# = snapFTENum("FTEHrsTot")
        Else
            FTEHrs# = 0
        End If
        snapBudget("JG_FTENUMFILL") = FTENum#
        snapBudget("JG_FTENUMVACN") = snapBudget("JG_FTENUM") - FTENum#
        snapBudget("JG_FTETOTHR") = FTEHrs#
        snapBudget.Update
    Else
        snapBudget("JG_FTENUMFILL") = 0
        snapBudget("JG_FTENUMVACN") = snapBudget("JG_FTENUM")
        snapBudget.Update
    End If
    snapFTENum.Close
    
    
    snapBudget.MoveNext
Loop
snapBudget.Close

MDIMain.panHelp(0).FloodPercent = 0
MDIMain.panHelp(0).ForeColor = &H0&
MDIMain.panHelp(0).FloodType = 0
mod_Upd_Pos_Budget_WFC = True

Exit Function


mod_Upd_Pos_Budget_WFC_Err:
If Err = 94 Then
Err = 0
Resume Next
End If
glbFrmCaption$ = "Module - Count Budgeted Positions"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update HRJOBBUD Count", "HRJOBBUD", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Public Function GetEmpSalary(EmpNbr)
Dim SQLQ
Dim rsSal As New ADODB.Recordset
Dim retVal
    retVal = 0
    
    SQLQ = "SELECT * "
    'If glbtermopen Then
    '    SQLQ = SQLQ & " from Term_SALARY_HISTORY "
    '    SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND TERM_SEQ=" & EmpNbr
    '    rsSal.Open SQLQ, gdbAdoIhr001X, adOpenStatic
    'Else
        SQLQ = SQLQ & " from HR_SALARY_HISTORY "
        SQLQ = SQLQ & " WHERE SH_CURRENT<>0 AND SH_EMPNBR=" & EmpNbr
        rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'End If
    
    If Not rsSal.EOF Then
        If Not IsNull(rsSal("SH_SALARY")) Then
            retVal = rsSal("SH_SALARY")
        End If
    End If
    rsSal.Close
    
    GetEmpSalary = retVal
End Function

Public Function getEmpPostion(xEmpNo)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT JH_EMPNBR,JH_CURRENT,JH_JOB FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) AND JH_EMPNBR = " & xEmpNo
    rs.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rs.EOF Then
        retVal = rs("JH_JOB")
    End If
    rs.Close
    getEmpPostion = retVal
End Function
Public Function getWFCPosFromJobSec(xAJob, xADiv)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    If Not IsNull(xAJob) And Not IsNull(xADiv) Then
        SQLQ = "SELECT JB_CODE, JB_DIV, JB_JOBCODE FROM HRJOB WHERE JB_JOBCODE ='" & xAJob & "' "
        SQLQ = SQLQ & "AND JB_DIV = '" & xADiv & "' "
        If rs.State <> 0 Then rs.Close
        rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rs.EOF Then
            If Not IsNull(rs("JB_CODE")) Then
                retVal = rs("JB_CODE")
            End If
        End If
        rs.Close
    End If
    getWFCPosFromJobSec = retVal
End Function
Public Sub CheckHRTABLCode(xName, xKey, xKeyDesc) 'Ticket #27742 Franks 11/10/2015
Dim rsTABL As New ADODB.Recordset
Dim SQLQ
    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & xName & "' AND TB_KEY = '" & xKey & "' "
    rsTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsTABL.EOF Then
        rsTABL.AddNew
        rsTABL("TB_COMPNO") = "001"
        rsTABL("TB_NAME") = xName
        rsTABL("TB_KEY") = xKey
        rsTABL("TB_DESC") = Trim(xKeyDesc)
        rsTABL("TB_LDATE") = Format(Now, "Short Date")
        rsTABL("TB_LTIME") = Time$
        rsTABL("TB_LUSER") = "999999999"
        rsTABL.Update
    End If
    rsTABL.Close
End Sub

Public Sub WFCNextPosNoSetup(xType) 'Ticket #27827 Franks 12/02/2015
Dim RsHRPARCO As New ADODB.Recordset
Dim SQLQ
    SQLQ = "SELECT * FROM HRPARCO "
    RsHRPARCO.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not RsHRPARCO.EOF Then
        If xType = "Reset" Then
            If IsNull(RsHRPARCO("PC_NEXT_POS_NBR")) Then
                    RsHRPARCO("PC_NEXT_POS_NBR") = WFCNextPosNoGen + 1
            Else
                'If RsHRPARCO("PC_NEXT_POS_NBR") = 0 Then
                    RsHRPARCO("PC_NEXT_POS_NBR") = WFCNextPosNoGen + 1
                'End If
            End If
            RsHRPARCO.Update
        Else 'non "Reset", Ongoing
            RsHRPARCO("PC_NEXT_POS_NBR") = RsHRPARCO("PC_NEXT_POS_NBR") + 1
            RsHRPARCO.Update
        End If
    End If
    RsHRPARCO.Close
End Sub
Private Function WFCNextPosNoGen() 'Ticket #27827 Franks 12/02/2015
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
Dim retVal
    retVal = 0
    SQLQ = "SELECT COUNT(JB_CODE) AS TOTNUM FROM HRJOB WHERE (1=1) " ' LEN(JB_CODE) > 6  "
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        retVal = rs("TOTNUM")
    End If
    rs.Close
    'retval = I ' xDiv & xPosGrp & Right("00000" & Trim(Str(I + 1)), 5)
    WFCNextPosNoGen = retVal
End Function

Public Function getNewPosCode(xDIV, xPosGrp) 'Ticket #27827 Franks 12/02/2015
Dim RsHRPARCO As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
Dim retVal
    retVal = ""
    I = 0
    'SQLQ = "SELECT COUNT(JB_CODE) AS TOTNUM FROM HRJOB WHERE  LEN(JB_CODE) > 6  "
    'rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'If Not rs.EOF Then
    '    I = rs("TOTNUM")
    'End If
    SQLQ = "SELECT * FROM HRPARCO "
    RsHRPARCO.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
    If Not RsHRPARCO.EOF Then
        I = RsHRPARCO("PC_NEXT_POS_NBR")
    End If
    RsHRPARCO.Close
    retVal = xDIV & xPosGrp & Right("00000" & Trim(Str(I)), 5)
    getNewPosCode = retVal
End Function

Public Sub WFCAutoPerformance(xEmpNo) 'Ticket #27774 Franks 12/30/2015
Dim rsEmp As New ADODB.Recordset
Dim rsPos As New ADODB.Recordset
Dim rsSal As New ADODB.Recordset
Dim rsPerf As New ADODB.Recordset
Dim rsFollow As New ADODB.Recordset
Dim SQLQ As String
Dim I As Integer
Dim xNDate, xDOH
'logic:
'"   On new hire, automatically create the HR_PERFORM_HISTORY with the Employee Number,  RA1 and Next Review Date = May 1st
'"   If the DOH's month is less than 5, the May 1st date will be the same year as the DOH
'"   If the DOH's month is 5 or greater, the May 1st date will be DOH year + 1.
'"   Only do if the employee's Union is EXEC or NONE.
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If rsEmp("ED_ORG") = "EXEC" Or rsEmp("ED_ORG") = "NONE" Then
            'open position record
            SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT JH_CURRENT = 0 AND JH_EMPNBR = " & xEmpNo & " "
            rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
            'open salary record
            SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & xEmpNo
            rsSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
            
            If Not rsPos.EOF And Not rsSal.EOF Then
                'Next Review Date = May 1st
                If Not IsNull(rsEmp("ED_DOH")) Then xDOH = rsEmp("ED_DOH") Else xDOH = Date
                xNDate = CVDate(("May 1, " & Year(xDOH)))
                If CVDate(xNDate) < CVDate(xDOH) Then xNDate = DateAdd("YYYY", 1, xNDate)
                
                SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_EMPNBR = " & xEmpNo & " "
                rsPerf.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                If rsPerf.EOF Then
                    rsPerf.AddNew
                    rsPerf("PH_EMPNBR") = xEmpNo
                    rsPerf("PH_CURRENT") = 1
                    rsPerf("PH_PREVIEW") = Date
                    rsPerf("PH_JOB") = rsPos("JH_JOB")
                    rsPerf("PH_REPTAU") = rsPos("JH_REPTAU")
                    rsPerf("PH_PNEXT") = xNDate
                    rsPerf("PH_JOB_ID") = rsPos("JH_ID")
                    rsPerf("PH_LDATE") = Date
                    rsPerf("PH_LTIME") = Time$
                    rsPerf("PH_LUSER") = glbUserID
                    rsPerf.Update
                    'follow up 'rsFollow
                    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & xEmpNo & " "
                    SQLQ = SQLQ & "AND EF_FREAS = 'PREV' "
                    SQLQ = SQLQ & "AND EF_FDATE = " & Date_SQL(xNDate) & " "
                    rsFollow.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic
                    If rsFollow.EOF Then
                        rsFollow.AddNew
                        rsFollow("EF_COMPNO") = "001"
                        rsFollow("EF_EMPNBR") = xEmpNo
                        rsFollow("EF_FDATE") = xNDate
                        rsFollow("EF_FREAS_TABL") = "FURE"
                        rsFollow("EF_FREAS") = "PREV"
                        rsFollow("EF_ADMINBY_TABL") = "EDAB"
                        rsFollow("EF_ADMINBY") = rsEmp("ED_ADMINBY")
                    End If
                    rsFollow("EF_LDATE") = Date
                    rsFollow("EF_LTIME") = Time$
                    rsFollow("EF_LUSER") = glbUserID
                    rsFollow.Update
                    rsFollow.Close
                End If
                rsPerf.Close
            End If
        End If
    End If
    rsEmp.Close
End Sub

Public Function IsMissingBudPos(xPos)
Dim rsPos As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = True
    If Not IsNull(xPos) Then
        SQLQ = "SELECT * FROM HRJOBBUD WHERE JG_CODE = '" & xPos & "' "
        SQLQ = SQLQ & "ORDER BY JG_FRDATE DESC"
        rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsPos.EOF Then
            retVal = True
        Else
            retVal = False
        End If
        rsPos.Close
    End If
    IsMissingBudPos = retVal
End Function

Public Function IsInactivePos(xPos)
Dim rsPos As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = True
    If Not IsNull(xPos) Then
        SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xPos & "' "
        rsPos.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsPos.EOF Then
            retVal = False
            If rsPos("JB_STATUS") = "INAC" Then
                retVal = True
            End If
            If Left(rsPos("JB_DESCR"), 2) = "Z " Then
                retVal = True
            End If
        End If
        rsPos.Close
    End If
    IsInactivePos = retVal
End Function

Public Function getTotHrsToESS(xEmpNo, xDATE, xReason) 'Ticket #28373 Franks 04/14/2016
Dim rsAtt As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = 0
    SQLQ = "SELECT SUM(AD_HRS) AS TOTHRS FROM HR_ATTENDANCE WHERE AD_EMPNBR =" & xEmpNo & " "
    SQLQ = SQLQ & "AND AD_REASON = '" & xReason & "' "
    SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xDATE) & " "
    SQLQ = SQLQ & "GROUP BY AD_EMPNBR "
    rsAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsAtt.EOF Then
        retVal = rsAtt("TOTHRS")
    End If
    rsAtt.Close
    getTotHrsToESS = retVal
End Function

Public Function getWFCPlantSecurity(xSection)
Dim SQLQ As String
Dim xStr
Dim retVal
    retVal = " (1=1) "
    If Len(glbPlantCode) > 0 Then
        If Not UCase(glbPlantCode) = "ALL" Then
            xStr = glbSeleSection
            xStr = Replace(xStr, "TB_KEY", "")
            xStr = Replace(xStr, "=", "")
            xStr = Replace(xStr, "'", "")
            retVal = "'" & xStr & "' LIKE '%'+" & xSection & "+'%' "
            'retval = " ('" & xSECTION & "' IN " & glbSeleSection & " ) "
        End If
    End If
    getWFCPlantSecurity = retVal
End Function

Public Function getWFCIfNetworkLoginDup(xEmpNo, xLogID)
Dim rsOther As New ADODB.Recordset
Dim xIsDuplicate As Boolean
Dim retVal
    
    xIsDuplicate = False
    
    'check active table
    SQLQ = "SELECT * FROM HREMP_OTHER WHERE NOT ER_EMPNBR = " & xEmpNo & " AND ER_NETWORKLOGIN = '" & xLogID & "' "
    rsOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsOther.EOF Then
        xIsDuplicate = True
    End If
    rsOther.Close
    If xIsDuplicate = False Then 'check Term table
        SQLQ = "SELECT * FROM Term_HREMP_OTHER WHERE ER_NETWORKLOGIN = '" & xLogID & "' "
        rsOther.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsOther.EOF Then
            xIsDuplicate = True
        End If
        rsOther.Close
    End If

    retVal = xIsDuplicate
    getWFCIfNetworkLoginDup = retVal
End Function

Public Function getWFCNetworkLoginNoDupicate(txtFName, txtSurname, txtMidName, xEmpNo, xLogID) 'Ticket #29246 Franks 11/02/2016
'"   Cannot create a duplicate.
'"   If a duplicate exists, use the first character of the middle name as the second character of the Network ID. Ie: James John Johnson & James Tom Johnson would have Network IDs of "jjjohnsn" and "jtjohnsn" respectively.
'"   If a duplicate still exist, don't use the Middle Initial. Add a numerical number to the last character of the Network ID. Start at zero and work up.
'"   Maximum number of characters is still 8.
Dim rsOther As New ADODB.Recordset
Dim xIsDuplicate As Boolean
Dim xlen As Integer
Dim xF
Dim xL
Dim I As Integer
Dim xStr
Dim xTmpID
Dim retVal
    
    retVal = xLogID
    xIsDuplicate = False

    'check duplicate - begin
    'check active table
    xIsDuplicate = getWFCIfNetworkLoginDup(xEmpNo, xLogID)
    If xIsDuplicate = False Then 'No duplicate found on both active and term, then use xLogID from getWFCNetworkLogin
        getWFCNetworkLoginNoDupicate = retVal
        Exit Function
    End If
    'check duplicate - end
    
    'Found duplicate ------------- begin
    
    'Use middle name
    '"   If a duplicate exists, use the first character of the middle name as the second character of the Network ID. Ie: James John Johnson & James Tom Johnson would have Network IDs of "jjjohnsn" and "jtjohnsn" respectively.
    If xIsDuplicate Then
        txtMidName = Trim(txtMidName)
        If Len(txtMidName) > 0 Then
            xlen = Len(xLogID)
            xF = Left(xLogID, 1)
            xL = Right(xLogID, xlen - 1)
            If Len(xL) >= 7 Then
                xL = Right(xL, 6)
            End If
            'xStr = Left(txtMidName, 1) & xL
            xTmpID = xF & Left(txtMidName, 1) & xL
            xTmpID = Left(Trim(xTmpID), 8)
            xTmpID = LCase(xTmpID)
            
            xIsDuplicate = getWFCIfNetworkLoginDup(xEmpNo, xTmpID)
            retVal = xTmpID
            If xIsDuplicate = False Then 'No duplicate found on both active and term, then use xLogID from getWFCNetworkLogin
                getWFCNetworkLoginNoDupicate = retVal
                Exit Function
            End If
    
        End If
    End If
    
    '"   If a duplicate still exist, don't use the Middle Initial. Add a numerical number to the last character of the Network ID
    If xIsDuplicate Then
        xTmpID = Left(xLogID, 7)
        'add single number
        For I = 1 To 9
            xTmpID = xTmpID & Trim(Str(I))
            xIsDuplicate = getWFCIfNetworkLoginDup(xEmpNo, xTmpID)
            retVal = xTmpID
            If xIsDuplicate = False Then 'No duplicate found on both active and term, then use xLogID from getWFCNetworkLogin
                getWFCNetworkLoginNoDupicate = retVal
                Exit Function
            End If
        Next
    End If
    
    'add 2 digits number
    If xIsDuplicate Then
        xTmpID = Left(xLogID, 6)
        'add single number
        For I = 10 To 99
            xTmpID = xTmpID & Trim(Str(I))
            xIsDuplicate = getWFCIfNetworkLoginDup(xEmpNo, xTmpID)
            retVal = xTmpID
            If xIsDuplicate = False Then 'No duplicate found on both active and term, then use xLogID from getWFCNetworkLogin
                getWFCNetworkLoginNoDupicate = retVal
                Exit Function
            End If
        Next
    End If
    
    'Found duplicate ------------- end
    
    getWFCNetworkLoginNoDupicate = retVal
End Function

Public Function getWFCNetworkLogin(txtFName, txtSurname) 'Ticket #28772 Franks 06/22/2016
Dim xTot As Integer
Dim I As Integer
Dim J As Integer
Dim xStr, xExcList
Dim xFName, xSurname
Dim retVal

'Login:
'On new hire, auto-create Network Login. The maximum login is 8 characters. The format is first initial of First Name and then if Surname is 7 characters or less,
'use the full Surname. Otherwise, drop the vowels starting from the right going back to the left.
'For example: Margaret Zyma's Network Login is "mzyma"; Darren Aspinall's Network Login is "daspinll"

    xFName = Trim(txtFName)
    xSurname = Trim(txtSurname)
    xSurname = Replace(xSurname, " ", "")
    If Len(xSurname) <= 7 Then
        retVal = Left(xFName, 1) & xSurname
    Else
        'drop the vowels starting from the right going back to the left.
        'You may have to hardcode the vowels into the program and if more than 8 characters delete one of  "a e i o u" Let's not include y in the program - Margaret
        xExcList = "aeio u"
        xTot = Len(xSurname)
        J = 1
        I = 1
        'For I = 1 To xTot
        Do While Len(xSurname) > 7 And I < xTot + 1
            I = I + 1
            'xStr = Right(xSurname, J)
            xStr = Mid(xSurname, Len(xSurname) - J + 1, 1)
            If InStr(xExcList, xStr) > 0 Then
                'xSurname = Left(xSurname, Len(xSurname) - 1) 'Mid(RHeading, 1, InStr(RHeading, "-"))
                'xSurname = Left(xSurname, Len(xSurname) - 2) & Right(xSurname, J - 1)
                xSurname = Left(xSurname, Len(xSurname) - J) & Right(xSurname, J - 1)
                J = J + 0
            Else
                J = J + 1
            End If
            If Len(xSurname) = 7 Then
                GoTo end_line
            End If
        Loop
        'Next
end_line:
        If Len(xSurname) > 7 Then
            xSurname = Left(xSurname, 7)
        End If
        retVal = Left(xFName, 1) & Left(xSurname, 7)
    End If
    getWFCNetworkLogin = LCase(retVal)
End Function

Public Function getWFCFiscalYearStartDate(xDATE)
Dim xYYYY, xMM
Dim retVal
    xYYYY = Year(xDATE)
    xMM = month(xDATE)
    'If xMM = 11 Or xMM = 12 Then
    If xMM < 11 Then
        xYYYY = xYYYY - 1
    End If
    retVal = CVDate("Nov 1, " & xYYYY)
    
    getWFCFiscalYearStartDate = retVal
End Function

Public Function getWFCFiscalYearToDate(xDATE)
Dim xYYYY, xMM
Dim retVal
    xYYYY = Year(xDATE)
    xMM = month(xDATE)
    If xMM = 11 Or xMM = 12 Then
        xYYYY = xYYYY + 1
    End If
    retVal = CVDate("Oct 31, " & xYYYY)
    
    getWFCFiscalYearToDate = retVal
End Function

Public Function getWFCAnnualSalary(xSal, xType, xWhrs)
Dim retVal
    retVal = 0
    If xType = "A" Then
        retVal = xSal
    End If
    If xType = "H" Then
        If IsNumeric(xWhrs) Then
            retVal = xSal * xWhrs * 52
        End If
    End If
    If xType = "M" Then
        retVal = xSal * 12
    End If
    If xType = "D" Then
        retVal = xSal * 5 * 52 'Assumes 5 day/week
    End If
    getWFCAnnualSalary = retVal
End Function

Public Function getRateToCAD(xYear, xMonth, xCURRENCYINDI, xCONVERT_NO)
Dim rsRate As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = 0
    If xCONVERT_NO = 1 Then 'CAD
        If Not IsNull(xCURRENCYINDI) Then
            If xCURRENCYINDI = "CAD" Then
                retVal = 1
            Else
                SQLQ = "SELECT * FROM HRIP_CURRENCY_EXCHG WHERE IP_YEAR = " & xYear & " AND IP_MTH_SEQ = '" & xMonth & "' "
                SQLQ = SQLQ & "AND IP_CURRENCYINDF = '" & xCURRENCYINDI & "' "
                SQLQ = SQLQ & "AND IP_CONVERT_NO = '" & xCONVERT_NO & "' "
                rsRate.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsRate.EOF Then '
                    If Not IsNull(rsRate("IP_RATE")) Then
                        retVal = rsRate("IP_RATE")
                    End If
                End If
                rsRate.Close
            End If
        End If
    End If
    If xCONVERT_NO = 2 Then 'USD
        If Not IsNull(xCURRENCYINDI) Then
            If xCURRENCYINDI = "USD" Then
                retVal = 1
            Else
                SQLQ = "SELECT * FROM HRIP_CURRENCY_EXCHG WHERE IP_YEAR = " & xYear & " AND IP_MTH_SEQ = '" & xMonth & "' "
                SQLQ = SQLQ & "AND IP_CURRENCYINDF = '" & xCURRENCYINDI & "' "
                SQLQ = SQLQ & "AND IP_CONVERT_NO = '" & xCONVERT_NO & "' "
                rsRate.Open SQLQ, gdbAdoIhr001, adOpenStatic
                If Not rsRate.EOF Then '
                    If Not IsNull(rsRate("IP_RATE")) Then
                        retVal = rsRate("IP_RATE")
                    End If
                End If
                rsRate.Close
            End If
        End If
    End If
    getRateToCAD = retVal
End Function

Private Function AUDITSALY_PUB(empNo, NSalary, oPayP, oJob, oGrid, oPayrollID, OSalCD, oHrsWk, OEDate, ONDate, OReason)
Dim TA As New ADODB.Recordset
Dim xADD As Boolean, xPT As String, xDIV As String
Dim TB As New ADODB.Recordset
Dim strFields As String

On Error GoTo AUDIT_ERR

AUDITSALY_PUB = False


TB.Open "SELECT ED_PT,ED_DIV FROM HREMP WHERE ED_EMPNBR = " & empNo, gdbAdoIhr001, adOpenForwardOnly
If Not TB.EOF Then
    If IsNull(TB("ED_PT")) Then
        xPT = ""
    Else
        xPT = TB("ED_PT")
    End If
    If IsNull(TB("ED_DIV")) Then
        xDIV = ""
    Else
        xDIV = TB("ED_DIV")
    End If
Else
    xPT = ""
    xDIV = ""
End If
TB.Close
'TA.Open "HRAUDIT", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic, adCmdTableDirect
'strfields added by Bryan 02/Dec/05 TICKET#9899
strFields = "AU_LOC_TABL, AU_EMP_TABL, AU_SUPCODE_TABL, AU_ORG_TABL, AU_PAYP_TABL, AU_BCODE_TABL, AU_TREAS_TABL,AU_DOLENT_TABL, "
strFields = strFields & "AU_EARN_TABL, AU_NEWEMP, AU_PTUPL, AU_DIVUPL, AU_SALARY, AU_OLDSAL, AU_PAYP, AU_OLDPAYP, AU_PAYP, "
strFields = strFields & "AU_OLDPAYP, AU_OLDPAYP, AU_JOB, AU_GRID, AU_PAYROLL_ID, AU_SALCD, AU_WHRS, AU_SEDATE, AU_SNDATE, "
strFields = strFields & "AU_COMPNO, AU_EMPNBR, AU_LDATE, AU_LUSER, AU_LTIME, AU_UPLOAD, AU_TYPE, AU_SREASON "
TA.Open "SELECT " & strFields & " FROM HRAUDIT WHERE 1=2", gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
xADD = False

'If OSalary <> NSalary Then GoTo MODUPD
''If OPayp <> NPayp Then GoTo MODUPD      'laura jan 28, 1998
'If OEDate <> NEDate Then GoTo MODUPD
'If ONDate <> NNDate Then GoTo MODUPD

'GoTo MODNOUPD

MODUPD:
TA.AddNew
TA("AU_LOC_TABL") = "EDLC": TA("AU_EMP_TABL") = "EDEM": TA("AU_SUPCODE_TABL") = "EDSP": TA("AU_ORG_TABL") = "EDOR"
TA("AU_PAYP_TABL") = "SDPP": TA("AU_BCODE_TABL") = "BNCD": TA("AU_TREAS_TABL") = "TERM": TA("AU_DOLENT_TABL") = "EDOL"
TA("AU_EARN_TABL") = "EARN"
TA("AU_NEWEMP") = "N"
TA("AU_PTUPL") = xPT
TA("AU_DIVUPL") = xDIV
TA("AU_SALARY") = NSalary
'TA("AU_OLDSAL") = NSalary  'Ticket #27056 - Do not save this as it will not export to Payroll if Salary = OldSal
TA("AU_PAYP") = oPayP ' FRANK 4/5/2000    'NPayp  Laura jan 28, 1998
TA("AU_OLDPAYP") = oPayP    '    ""
TA("AU_JOB") = oJob          ' FRANK 4/5/2000
TA("AU_GRID") = oGrid
If glbMulti Then TA("AU_PAYROLL_ID") = oPayrollID
TA("AU_SALCD") = OSalCD
TA("AU_WHRS") = oHrsWk 'ADDED BY RAUBREY 7/7/97
'If OEDate <> NEDate Then TA("AU_SEDATE") = IIf(IsDate(NEDate), NEDate, Null)   'Jaddy 11/15/99
'If ONDate <> NNDate Then TA("AU_SNDATE") = IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99
TA("AU_SEDATE") = OEDate   'IIf(IsDate(NEDate), NEDate, Null)   'Jaddy 11/15/99
TA("AU_SNDATE") = ONDate   'IIf(IsDate(NNDate), NNDate, Null)  'Jaddy 11/15/99

'Ticket #23666 - Update with Salary Reason for Change as well.
TA("AU_SREASON") = OReason

TA("AU_COMPNO") = "001"
TA("AU_EMPNBR") = empNo

'Ticket #23943 - Town of Orangeville noticed the LDATE was not getting updated properly - Jerry asked to fix this as per Salary screen.
If glbCompSerial = "S/N - 2227W" And (xPT = "SE" Or xPT = "OT") Then ' CCAC Kingston, see ticket #3296
    'TA("AU_LDATE") = Format(DateAdd("d", 14, NEDate), "SHORT DATE")
    TA("AU_LDATE") = Format(DateAdd("d", 14, OEDate), "SHORT DATE")
Else
    'Ticket #23943 - Town of Orangeville
    If glbCompSerial = "S/N - 2383W" Then
        'If CVDate(NEDate) > CVDate(Date) Then
        If CVDate(OEDate) > CVDate(Date) Then
            'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
            TA("AU_LDATE") = Format(OEDate, "SHORT DATE")
        Else
            TA("AU_LDATE") = Date
        End If
    Else
        'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")
        TA("AU_LDATE") = Format(OEDate, "SHORT DATE")
    End If
End If
'TA("AU_LDATE") = Format(NEDate, "SHORT DATE")

TA("AU_LUSER") = glbUserID
TA("AU_LTIME") = Time$
TA("AU_UPLOAD") = "N"
TA("AU_TYPE") = "A"
'If glbSoroc Or glbSyndesis Then
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ
    SQLQ = "SELECT ED_EMPNBR,ED_PAYROLL_ID FROM HREMP WHERE ED_EMPNBR =" & empNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        If Not IsNull(rsEmp("ED_PAYROLL_ID")) Then TA("AU_PAYROLL_ID") = rsEmp("ED_PAYROLL_ID")
    End If
    rsEmp.Close
'End If
TA.Update


MODNOUPD:
AUDITSALY_PUB = True
Exit Function
AUDIT_ERR:

'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "ADDING AUDIT RECORD", "AUDIT FILE", "SAME SALARY UPDATE")
'If gintRollBack% = False Then Resume Next Else Unload Me
MsgBox Err.Description
End Function

Public Sub WFCPosReptsUpd(xEmpNo, xOldReptEmpNo, xNewReptEmpNo, xEffDate) 'Ticket #29343 Franks 10/24/2016
'see modUpdateSelection_On_LAYOFF
Dim rsJOB As New ADODB.Recordset
Dim SQLQ As String
Dim oPHRS, oWHRS, ODHRS, oJob As String, OSDATE
Dim OLeadHand, OLabourCD, OReason
Dim oOrg, oDeptNo, oStatus, oGLNo, oComment, oComment2
Dim oPayCategory
Dim OLambtonJob
Dim OFTE, fOldFTE, fNewFTE, fFTEDate
Dim oLABOUREDATE
Dim oENDDATE, oEndReason
Dim OBillingRate
Dim oSHIFT As String, oREPTAU As String
Dim nJobID
Dim oRepAut, oRepAut2, oRepAut3, oRepAut4
Dim oRepA1EDate, oRepA2EDate, oRepA3EDate, oRepA4EDate
Dim oFTENum, oFTEHrs
Dim oPTFT
Dim oDiv, oDept, oEmp, oRegion, oSect, oPosCtrl, oPayCateg, oBillRate, oLoc, oPosStatus
Dim xNewEffDate

    If Not IsNumeric(xOldReptEmpNo) Then Exit Sub
    If xOldReptEmpNo = xNewReptEmpNo Then Exit Sub 'No Rept change then skip
    
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT JH_CURRENT = 0 AND JH_EMPNBR = " & xEmpNo & " "
    SQLQ = SQLQ & "AND (JH_REPTAU = " & xOldReptEmpNo & " OR JH_REPTAU2 = " & xOldReptEmpNo & " OR JH_REPTAU3 = " & xOldReptEmpNo & " OR JH_REPTAU4 = " & xOldReptEmpNo & ")"
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsJOB.EOF Then
        Exit Sub
    End If
    
    'Clear fields
    oJob = "": OReason = "": OSDATE = "": ODHRS = "": oWHRS = "": oPHRS = "": oRepAut = "": oSHIFT = "": oFTENum = "": oFTEHrs = ""
    oOrg = "": oPTFT = "": oComment = "": oComment2 = "": oRepAut2 = "": oRepAut3 = "": oRepAut4 = "": oDiv = "": oDeptNo = "": oEmp = ""
    oGLNo = "": oSect = "": oRegion = "": oPosCtrl = "": oPayrollID = "": oGrid = "": oPayCateg = "": oBillRate = ""
    oRepA1EDate = "": oRepA2EDate = "": oRepA3EDate = "": oRepA4EDate = ""
    oPosStatus = ""
    If Not rsJOB.EOF Then
        'Retrieve existing Current Position Data
        oJob = rsJOB("JH_JOB")
        OReason = rsJOB("JH_JREASON")
        OSDATE = rsJOB("JH_SDATE")
        xNewEffDate = xEffDate
        If IsDate(OSDATE) And IsDate(xNewEffDate) Then
            If CVDate(OSDATE) > CVDate(xNewEffDate) Then
                xNewEffDate = OSDATE
            End If
        End If
        If Not IsNull(rsJOB("JH_DHRS")) Then ODHRS = rsJOB("JH_DHRS")
        If Not IsNull(rsJOB("JH_WHRS")) Then oWHRS = rsJOB("JH_WHRS")
        If Not IsNull(rsJOB("JH_PHRS")) Then oPHRS = rsJOB("JH_PHRS")
        If Not IsNull(rsJOB("JH_SHIFT")) Then oSHIFT = rsJOB("JH_SHIFT")
        If Not IsNull(rsJOB("JH_FTENUM")) Then oFTENum = rsJOB("JH_FTENUM")
        If Not IsNull(rsJOB("JH_FTEHRS")) Then oFTEHrs = rsJOB("JH_FTEHRS")
        If Not IsNull(rsJOB("JH_ORG")) Then oOrg = rsJOB("JH_ORG")
        If Not IsNull(rsJOB("JH_PT")) Then oPTFT = rsJOB("JH_PT")
        If Not IsNull(rsJOB("JH_COMMENT")) Then oComment = rsJOB("JH_COMMENT")
        If Not IsNull(rsJOB("JH_COMMENT2")) Then oComment2 = rsJOB("JH_COMMENT2")
        If Not IsNull(rsJOB("JH_REPTAU")) Then oRepAut = rsJOB("JH_REPTAU")
        If Not IsNull(rsJOB("JH_REPTAU2")) Then oRepAut2 = rsJOB("JH_REPTAU2")
        If Not IsNull(rsJOB("JH_REPTAU3")) Then oRepAut3 = rsJOB("JH_REPTAU3")
        If Not IsNull(rsJOB("JH_REPTAU4")) Then oRepAut4 = rsJOB("JH_REPTAU4")
        If Not IsNull(rsJOB("JH_EDATEREPT1")) Then oRepA1EDate = rsJOB("JH_EDATEREPT1")
        If Not IsNull(rsJOB("JH_EDATEREPT2")) Then oRepA2EDate = rsJOB("JH_EDATEREPT2")
        If Not IsNull(rsJOB("JH_EDATEREPT3")) Then oRepA3EDate = rsJOB("JH_EDATEREPT3")
        If Not IsNull(rsJOB("JH_EDATEREPT4")) Then oRepA4EDate = rsJOB("JH_EDATEREPT4")
        If Not IsNull(rsJOB("JH_PAYROLL_ID")) Then oPayrollID = rsJOB("JH_PAYROLL_ID")
        ''oDiv = rsJOB("JH_DIV")
        ''oDeptNo = rsJOB("JH_DEPTNO")
        ''oEmp = rsJOB("JH_EMP")
        ''oGLNo = rsJOB("JH_GLNO")
        ''oSect = rsJOB("JH_SECTION")
        ''oRegion = rsJOB("JH_REGION")
        ''oPosCtrl = rsJOB("JH_POSITION_CONTROL")
        ''oGrid = rsJOB("JH_GRID")
        ''oPayCateg = rsJOB("JH_PAYROLL_CATEGORY")
        ''oBillRate = rsJOB("JH_BILLINGRATE")
        '''oENDDATE = rsJOB("JH_ENDDATE")
        '''oEndReason = rsJOB("JH_ENDREAS")
        ''oPosStatus = rsJOB("JH_ESTATUS")
        lngLastCurrentID& = rsJOB("JH_ID")

        'Remove the Current check from the existing current position record
        rsJOB("JH_CURRENT") = False
        rsJOB("JH_ENDDATE") = DateAdd("d", -1, CVDate(xNewEffDate))
        'rsJOB("JH_ENDREAS") = "" 'clpNReason.Text
        rsJOB("JH_LDATE") = Format(Now, "Short Date")
        rsJOB("JH_LTIME") = Time$
        rsJOB("JH_LUSER") = glbUserID
        rsJOB.Update
        
        'Add a new Current Position records
        rsJOB.AddNew
        rsJOB("JH_COMPNO") = "001"
        rsJOB("JH_EMPNBR") = xEmpNo 'empNo&
        rsJOB("JH_JOB") = oJob
        rsJOB("JH_SDATE") = CVDate(xNewEffDate) ' OSDATE
        If ODHRS <> "" And Not IsNull(ODHRS) Then rsJOB("JH_DHRS") = ODHRS
        If oWHRS <> "" And Not IsNull(oWHRS) Then rsJOB("JH_WHRS") = oWHRS
        If oPHRS <> "" And Not IsNull(oPHRS) Then rsJOB("JH_PHRS") = oPHRS
        rsJOB("JH_JREASON") = "TITL" 'OReason
        rsJOB("JH_CURRENT") = True
        
        If oSHIFT <> "" And Not IsNull(oSHIFT) Then rsJOB("JH_SHIFT") = oSHIFT
        If oFTENum <> "" And Not IsNull(oFTENum) Then rsJOB("JH_FTENUM") = oFTENum
        If oFTEHrs <> "" And Not IsNull(oFTEHrs) Then rsJOB("JH_FTEHRS") = oFTEHrs
        
        If oOrg <> "" And Not IsNull(oOrg) Then rsJOB("JH_ORG") = oOrg
        If oPTFT <> "" And Not IsNull(oPTFT) Then rsJOB("JH_PT") = oPTFT
        If oComment <> "" And Not IsNull(oComment) Then rsJOB("JH_COMMENT") = oComment
        If oComment2 <> "" And Not IsNull(oComment2) Then rsJOB("JH_COMMENT2") = oComment2
        
        '----- rept change logic ---------------------- begin
        If oRepAut = xOldReptEmpNo And Not (xOldReptEmpNo = xNewReptEmpNo) Then
            If IsNumeric(xNewReptEmpNo) Then
                rsJOB("JH_REPTAU") = xNewReptEmpNo
                rsJOB("JH_EDATEREPT1") = CVDate(xNewEffDate)
            Else
                rsJOB("JH_EDATEREPT1") = Null
            End If
            'rsJOB("JH_EDATEREPT1") = CVDate(xNewEffDate)
        Else
            If IsDate(oRepA1EDate) Then rsJOB("JH_EDATEREPT1") = CVDate(oRepA1EDate)
            If oRepAut <> "" And Not IsNull(oRepAut) Then
                rsJOB("JH_REPTAU") = oRepAut
                If IsNull(rsJOB("JH_EDATEREPT1")) Then
                    rsJOB("JH_EDATEREPT1") = CVDate(xNewEffDate)
                End If
            End If
        End If
        If oRepAut2 = xOldReptEmpNo And Not (xOldReptEmpNo = xNewReptEmpNo) Then
            If IsNumeric(xNewReptEmpNo) Then
                rsJOB("JH_REPTAU2") = xNewReptEmpNo
                rsJOB("JH_EDATEREPT2") = CVDate(xNewEffDate)
            Else
                rsJOB("JH_EDATEREPT2") = Null
            End If
            'rsJOB("JH_EDATEREPT2") = CVDate(xNewEffDate)
        Else
            If oRepAut2 <> "" And Not IsNull(oRepAut2) Then rsJOB("JH_REPTAU2") = oRepAut2
            If IsDate(oRepA2EDate) Then rsJOB("JH_EDATEREPT2") = CVDate(oRepA2EDate)
        End If
        
        If oRepAut3 = xOldReptEmpNo And Not (xOldReptEmpNo = xNewReptEmpNo) Then
            If IsNumeric(xNewReptEmpNo) Then
                rsJOB("JH_REPTAU3") = xNewReptEmpNo
                rsJOB("JH_EDATEREPT3") = CVDate(xNewEffDate)
            Else
                rsJOB("JH_EDATEREPT3") = Null
            End If
            'rsJOB("JH_EDATEREPT3") = CVDate(xNewEffDate)
        Else
            If oRepAut3 <> "" And Not IsNull(oRepAut3) Then rsJOB("JH_REPTAU3") = oRepAut3
            If IsDate(oRepA3EDate) Then rsJOB("JH_EDATEREPT3") = CVDate(oRepA3EDate)
        End If
        If oRepAut4 = xOldReptEmpNo And Not (xOldReptEmpNo = xNewReptEmpNo) Then
            If IsNumeric(xNewReptEmpNo) Then
                rsJOB("JH_REPTAU4") = xNewReptEmpNo
                rsJOB("JH_EDATEREPT4") = CVDate(xNewEffDate)
            Else
                rsJOB("JH_EDATEREPT4") = Null
            End If
            'rsJOB("JH_EDATEREPT4") = CVDate(xNewEffDate)
        Else
            If oRepAut4 <> "" And Not IsNull(oRepAut4) Then rsJOB("JH_REPTAU4") = oRepAut4
            If IsDate(oRepA4EDate) Then rsJOB("JH_EDATEREPT4") = CVDate(oRepA4EDate)
        End If
        '----- rept change logic ---------------------- end
        
        ''rsJOB("JH_DIV") = clpNDiv.Text
        ''rsJOB("JH_DEPTNO") = clpNDept.Text
        ''rsJOB("JH_EMP") = oEmp
        ''rsJOB("JH_GLNO") = oGLNo
        ''rsJOB("JH_SECTION") = oSect
        ''rsJOB("JH_REGION") = oRegion
        ''rsJOB("JH_POSITION_CONTROL") = oPosCtrl
        ''rsJOB("JH_PAYROLL_ID") = oPayrollID
        ''rsJOB("JH_GRID") = oGrid
        ''rsJOB("JH_PAYROLL_CATEGORY") = oPayCateg
        ''rsJOB("JH_BILLINGRATE") = oBillRate
        'rsJOB("JH_ENDDATE") = rEndDate
        'rsJOB("JH_ENDREAS") = rEndReason
        If oPosStatus <> "" And Not IsNull(oPosStatus) Then rsJOB("JH_ESTATUS") = oPosStatus
        
        rsJOB("JH_LDATE") = Format(Now, "Short Date")
        rsJOB("JH_LTIME") = Time$
        rsJOB("JH_LUSER") = glbUserID
        rsJOB.Update
        
    End If
    rsJOB.Close
    
    
End Sub

Public Sub WFCPubPosChangedcmdEmail(xEmpNo, xMBody, xSubject) 'Ticket #29343 Franks 10/25/2016
Dim xEmail
Dim xToEmail As String
Dim xEmpName
Dim MailBody
    
On Error GoTo ErrorHandler

    MailBody = xMBody
    If Not gsEMAIL_ONPOSITION Then
        Exit Sub
    End If
    If Not UserEmailExist Then
        Exit Sub
    End If
        
    xToEmail = GetComPreferEmail("EMAIL_ONPOSITION", xEmpNo)
    If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
        xToEmail = GetComPreferEmail("EMAIL_ONPOSITION")
    End If
    xEmpName = GetEmpData(xEmpNo, "ED_SURNAME") & "," & GetEmpData(xEmpNo, "ED_FNAME")
    
    frmSendEmail.txtTo.Text = xToEmail
    'frmSendEmail.txtSubject.Text = "info:HR Position Change Notice - " & xEmpName ' lblEEName.Caption
    frmSendEmail.txtSubject.Text = xSubject & " - " & xEmpName ' lblEEName.Caption
    frmSendEmail.txtBody.Text = MailBody
    
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(0).Caption = "Sending email..."
    frmSendEmail.Tag = ""
    frmSendEmail.cmdSend_Click
    Do
        DoEvents
    Loop Until frmSendEmail.Tag <> ""   ' MC - dkostka - 05/03/01 - Changed from = "DONE" to <> ""
    ' AC - dkostka - 05/03/01 - Added checking to make sure the email went through,
    '   otherwise refuse to terminate the employee.
    If frmSendEmail.Tag = "DONE" Then
        Unload frmSendEmail
        AbortTerm = False
    Else
        Unload frmSendEmail
        AbortTerm = True
    End If
    MDIMain.panHelp(0).Caption = ""
    MDIMain.panHelp(0).FloodType = 1

exH:
    Exit Sub
    
ErrorHandler:
    If Err.Number = 364 Then Exit Sub
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error #" & Err.Number
    Resume exH
    
End Sub


Public Function IsWFCReptPosAuth(xPosCode) 'Ticket #29507 Franks 11/30/2016
Dim rs As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset
Dim rsJobHis As New ADODB.Recordset
Dim SQLQ As String
Dim xReptPosCode
Dim retVal As Boolean

    retVal = False
    SQLQ = "SELECT * FROM HRJOB WHERE (1=1) "
    SQLQ = SQLQ & "AND JB_REPTAU = '" & xPosCode & "' "
    'SQLQ = SQLQ & "ORDER BY ED_SURNAME, ED_FNAME"
    If rs.State <> 0 Then rs.Close
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rs.EOF Then
        IsWFCReptPosAuth = retVal
        Exit Function
    End If

    
    'populate the employees who report to this Position
    gdbAdoIhr001W.BeginTrans
    gdbAdoIhr001W.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP = '" & glbUserID & "' "
    gdbAdoIhr001W.CommitTrans
    
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    If rsEListWRK.State <> 0 Then rsEListWRK.Close
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    Do While Not rs.EOF
        xReptPosCode = rs("JB_CODE")
        
        SQLQ = "SELECT JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 AND JH_JOB = '" & xReptPosCode & "' "
        If rsJobHis.State <> 0 Then rsJobHis.Close
        rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
                
        Do While Not rsJobHis.EOF
            If IsNull(rsJobHis("JH_REPTAU")) Then 'Ticket #30465 Franks 08/03/2017
            Else
                rsEListWRK.AddNew
                rsEListWRK("TT_COMPNO") = "001"
                rsEListWRK("TT_EMPNBR") = rsJobHis("JH_EMPNBR")
                rsEListWRK("TT_SURNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_SURNAME")
                rsEListWRK("TT_FNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_FNAME")
                rsEListWRK("TT_WRKEMP") = glbUserID
                rsEListWRK.Update
                retVal = True
                xCunt = xCunt + 1
            End If
            
            rsJobHis.MoveNext
        Loop
        rs.MoveNext
    Loop
    rs.Close
    
    IsWFCReptPosAuth = retVal
    
End Function

Public Function IsWFCReptAuth(xEmpNo, xExcludeNo) 'Ticket #29438 Franks 11/08/2016
Dim rsEmp As New ADODB.Recordset
Dim SQLQ As String
Dim retVal As Boolean

    retVal = False
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR IN (SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE JH_CURRENT = 1 "
    SQLQ = SQLQ & "AND JH_REPTAU = " & xEmpNo & ") "
    If IsNumeric(xExcludeNo) Then
        SQLQ = SQLQ & "AND NOT ED_EMPNBR = " & xExcludeNo & " "  'Not include current emp
    End If
    SQLQ = SQLQ & "ORDER BY ED_SURNAME, ED_FNAME"
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmp.EOF Then
        retVal = True
    End If
    rsEmp.Close
    IsWFCReptAuth = retVal
End Function


Public Sub PubReptPosEmpListByEmp(xEmpNo)
Dim rsEmp As New ADODB.Recordset
Dim rsJOB As New ADODB.Recordset
Dim rsEListWRK As New ADODB.Recordset
Dim rsJobHis As New ADODB.Recordset
Dim Comp$, Job$, JobCount&, SQLQ As String, pct#, ipct#, rcount&, spct%
Dim JobPoints#
Dim snapEvalPoints As New ADODB.Recordset
Dim FTENum#, FTEHrs#
Dim snapFTENum As New ADODB.Recordset
Dim snapFTEHrs As New ADODB.Recordset
Dim snapBudget As New ADODB.Recordset
Dim xJob, xDIV, xDeptno, xGLNO, xPosCtrl
Dim xSec, xCunt
Dim xBudgNo, xVacantNo, I
Dim xReptPosCode, xReptPosDesc, xPosCodeDesc, xPosCode
Dim retVal
    'retVal = ""
    If Len(xEmpNo) = 0 Then
        'getReptPosEmpListByEmp = retVal
        Exit Sub
    End If
    If Not IsNumeric(xEmpNo) Then ' Len(xPosCode) = 0 Then
        'getReptPosEmpListByEmp = retVal
        Exit Sub
    End If
    
    'Exit Sub '
    
    gdbAdoIhr001.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 "
    If Len(xEmpNo) > 0 Then
        If IsNumeric(xEmpNo) Then
            SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
        End If
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If rsEmp.EOF Then
        'getReptPosEmpListByEmp = retVal
        Exit Sub
    End If
    If Not rsEmp.EOF Then
        xPosCode = rsEmp("JH_JOB")
    End If
    
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & xPosCode & "' "
    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
    'xReptPosCode = xPosCode '""
    xPosCodeDesc = ""
    If rsJOB.EOF Then
        'getReptPosEmpListByEmp = retVal
        Exit Sub
    Else
        xPosCodeDesc = rsJOB("JB_DESCR")
    End If
    
    xCunt = 0
    
    SQLQ = "SELECT JH_EMPNBR,JH_REPTAU FROM HR_JOB_HISTORY WHERE JH_CURRENT<>0 " 'AND JH_JOB = '" & xReptPosCode & "' "
    If Len(xEmpNo) > 0 Then
        If IsNumeric(xEmpNo) Then
            SQLQ = SQLQ & "AND JH_REPTAU = " & xEmpNo & " "
            'SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
        End If
    End If
    rsJobHis.Open SQLQ, gdbAdoIhr001, adOpenStatic
    Do While Not rsJobHis.EOF
        rsEListWRK.AddNew
        rsEListWRK("TT_COMPNO") = "001"
        rsEListWRK("TT_EMPNBR") = rsJobHis("JH_EMPNBR")
        rsEListWRK("TT_SURNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_SURNAME")
        rsEListWRK("TT_FNAME") = GetEmpData(rsJobHis("JH_EMPNBR"), "ED_FNAME")
        rsEListWRK("TT_WRKEMP") = glbUserID
        rsEListWRK.Update
        xCunt = xCunt + 1

        rsJobHis.MoveNext
    Loop
    rsJobHis.Close

    'getReptPosEmpListByEmp = retVal
 
    'Me.vbxCrystal2.ReportFileName = glbIHRREPORTS & "RZEmpList4.rpt" '"RZEmpList3.rpt"
    'Me.vbxCrystal2.SelectionFormula = " {HR_EMPLIST_WRK.TT_WRKEMP}='" & glbUserID & "'"
    'Me.vbxCrystal2.Formulas(0) = "rTitle='" & lblReptAuth(0).Caption & " information'"
    'Me.vbxCrystal2.Connect = RptODBC_SQL
    ''window title if appropriate
    'Me.vbxCrystal2.WindowTitle = lblReptAuth(0).Caption & " Employee Position Information"
    'Me.vbxCrystal2.Destination = 0
    'Screen.MousePointer = DEFAULT
    'Me.vbxCrystal2.Action = 1
    'vbxCrystal2.Reset


End Sub

Public Function getEmpcurrencyIndi(xEmpNo) 'Ticket #29886 Franks 02/27/2017
Dim rsEmpSal As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE NOT SH_CURRENT = 0 AND SH_EMPNBR = " & xEmpNo & " "
    rsEmpSal.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpSal.EOF Then
        If Not IsNull(rsEmpSal(("SH_CURRENCYINDI"))) Then
            retVal = Trim(rsEmpSal(("SH_CURRENCYINDI")))
        End If
    End If
    getEmpcurrencyIndi = retVal
End Function

Public Function WFCSigningApproLimitUpt(xYear, xPlant)
Dim rsPos As New ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim SQLQ As String
Dim xLimit, xBand, xMarketline
Dim I As Long
Dim retVal
    retVal = 0
    WFCSigningApproLimitUpt = retVal
    If Not IsNumeric(xYear) Then Exit Function
    If Len(xPlant) = 0 Then Exit Function
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_SECTION = '" & xPlant & "' "
    'SQLQ = SQLQ & "AND JB_BAND = '" & xBand & "' "
    'SQLQ = SQLQ & "AND JB_MARKETLINE = '" & xMarketline & "' "
    SQLQ = SQLQ & "AND NOT (JB_STATUS = 'INAC') "
    rsPos.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    Screen.MousePointer = HOURGLASS
    MDIMain.panHelp(0).FloodType = 1
    I = 0
    Do While Not rsPos.EOF
        MDIMain.panHelp(0).FloodPercent = (I / rsPos.RecordCount) * 100: I = I + 1
        DoEvents
        
        If Not IsNull(rsPos("JB_BAND")) Then xBand = rsPos("JB_BAND") Else xBand = ""
        If Not IsNull(rsPos("JB_MARKETLINE")) Then xMarketline = rsPos("JB_MARKETLINE") Else xMarketline = ""
        
        If Len(xBand) > 0 And Len(xMarketline) > 0 Then
            SQLQ = "SELECT * FROM WFC_Salary_Administration WHERE (1=1) "
            SQLQ = SQLQ & "AND SectionCode = '" & xPlant & "' "
            SQLQ = SQLQ & "AND FiscalYear = " & xYear & " "
            SQLQ = SQLQ & "AND BAND = '" & xBand & "' "
            SQLQ = SQLQ & "AND MarketLine = '" & xMarketline & "' "
            If rsGrid.State <> 0 Then rsGrid.Close
            rsGrid.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsGrid.EOF Then
                If Not IsNull(rsGrid("APPR_LIMIT")) Then
                    If IsNumeric(rsGrid("APPR_LIMIT")) Then
                        rsPos("JB_APPR_LIMIT") = rsGrid("APPR_LIMIT") ' xLimit
                        rsPos.Update
                        retVal = retVal + 1
                    End If
                End If
            End If
            'rsPos.MoveNext
            'retval = retval + 1
        End If
        
        rsPos.MoveNext
    Loop
    rsPos.Close
    
    MDIMain.panHelp(0).FloodType = 0
    Screen.MousePointer = DEFAULT
        
    WFCSigningApproLimitUpt = retVal
End Function

Public Function WFCSigningApprovalLimitUpt(xPlant, xBand, xMarketline, xLimit)  'Ticket #29846 Franks 03/06/2017
Dim rsPos As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = 0
    WFCSigningApprovalLimitUpt = retVal
    
    If Not IsNumeric(xLimit) Then Exit Function
    If Len(xPlant) = 0 Then Exit Function
    If Len(xBand) = 0 Then Exit Function
    If Len(xMarketline) = 0 Then Exit Function
    
    SQLQ = "SELECT * FROM HRJOB WHERE JB_SECTION = '" & xPlant & "' "
    SQLQ = SQLQ & "AND JB_BAND = '" & xBand & "' "
    SQLQ = SQLQ & "AND JB_MARKETLINE = '" & xMarketline & "' "
    SQLQ = SQLQ & "AND NOT (JB_STATUS = 'INAC') "
    rsPos.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsPos.EOF
        rsPos("JB_APPR_LIMIT") = xLimit
        rsPos.Update
        rsPos.MoveNext
        retVal = retVal + 1
    Loop
    rsPos.Close
    WFCSigningApprovalLimitUpt = retVal
End Function

Public Function WFCSigningApprovalLimitGet(xPlant, xBand, xMarketline, Optional xYear)  'Ticket #29846 Franks 03/06/2017
Dim rsPos As New ADODB.Recordset
Dim SQLQ As String
Dim xLocYear
Dim retVal
    retVal = 0
    WFCSigningApprovalLimitGet = retVal
    
    xLocYear = ""
    If Not IsMissing(xYear) Then
        If Len(xYear) > 0 Then
            xLocYear = xYear
        End If
    End If
    
    If IsNull(xBand) Then Exit Function
    If IsNull(xMarketline) Then Exit Function
    If Len(xPlant) = 0 Then Exit Function
    If Len(xBand) = 0 Then Exit Function
    If Len(xMarketline) = 0 Then Exit Function
    
    SQLQ = "SELECT * FROM WFC_Salary_Administration WHERE SectionCode = '" & xPlant & "' "
    SQLQ = SQLQ & "AND BAND = '" & xBand & "' "
    SQLQ = SQLQ & "AND MarketLine = '" & xMarketline & "' "
    If Len(xLocYear) > 0 Then
        If IsNumeric(xLocYear) Then
            SQLQ = SQLQ & "AND FiscalYear = " & xLocYear & " "
        End If
    End If
    SQLQ = SQLQ & "ORDER BY FiscalYear DESC "
    rsPos.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsPos.EOF Then
        If Not IsNull(rsPos("APPR_LIMIT")) Then
            retVal = rsPos("APPR_LIMIT")
        End If
    End If
    rsPos.Close
    WFCSigningApprovalLimitGet = retVal
End Function

Public Function getDivField(xDIV, Field) 'Ticket #30012 Franks 04/07/2017
Dim rsDiv As New ADODB.Recordset
Dim SQLQ As String
Dim retVal
    retVal = ""
    rsDiv.Open "SELECT DIV," & Field & " FROM HR_DIVISION WHERE DIV = '" & xDIV & "' ", gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsDiv.EOF Then
        If Not IsNull(rsDiv(Field)) Then
            retVal = rsDiv(Field)
        End If
    End If
    rsDiv.Close
    getDivField = retVal
End Function

Public Function getWFCCurrencyIndi(xPlantCode)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ, xStr
Dim retVal

    retVal = ""
    If Len(xPlantCode) > 0 Then
        SQLQ = "select * from WFC_Salary_Administration "
        SQLQ = SQLQ & " WHERE SectionCode ='" & xPlantCode & "' "
        SQLQ = SQLQ & " AND NOT ( CurrencyIndicator IS NULL OR CurrencyIndicator = '') "
        SQLQ = SQLQ & "ORDER BY FiscalYear DESC"
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp("CurrencyIndicator")) Then
                retVal = rsTemp("CurrencyIndicator")
            End If
        End If
        rsTemp.Close
    End If
    getWFCCurrencyIndi = retVal
End Function

Public Function getWFC_CONP_Pos(xDIV)
Dim rsTemp As New ADODB.Recordset
Dim SQLQ, xStr
Dim retVal
    retVal = ""
    If Len(xDIV) > 0 Then
        SQLQ = "SELECT * FROM HRJOB WHERE LEFT(JB_CODE,7)= '" & xDIV & "IND' "
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            retVal = rsTemp("JB_CODE")
        End If
        rsTemp.Close
    End If
    getWFC_CONP_Pos = retVal
End Function
