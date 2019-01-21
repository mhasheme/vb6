Attribute VB_Name = "EmpTerm"
Dim db001 As Database

Function TERM_ATTENDANCE(EEID As Long, DtTm)
Dim SQLQ As String, dtATTD%, SQLQ2 As String
Dim iRow As Integer, Msg As String
Dim YLen
TERM_ATTENDANCE = False
On Error GoTo TERM_ATTENDANCE_Err

SQLQ = "INSERT INTO Term_ATTENDANCE "
' danielk - 04/11/2003 - was removing some things on termination, added COMM through POINT
If glbchkSum = True Then
    SQLQ = SQLQ & "(AD_COMPNO,AD_EMPNBR,AD_REASON,AD_DOA,AD_HRS,AD_POINT,AD_LDATE,AD_LTIME,AD_LUSER,TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT AD_COMPNO, AD_EMPNBR, AD_REASON, "
    'date
    If glbOracle Then
        SQLQ = SQLQ & " TO_DATE('31-12-' || TO_CHAR(AD_DOA,'YYYY'),'DD-MM-YYYY') AS AD_DOA,"
    ElseIf glbSQL Then
        SQLQ = SQLQ & " '" & IIf(glbFrench, "déc", "Dec") & " 31,'+CONVERT(varchar(4),AVG(Year(AD_DOA))) AS AD_DOA,"
    Else
        SQLQ = SQLQ & "CVDATE('" & IIf(glbFrench, "déc", "Dec") & " 31,'+STR(Year(AD_DOA))) As AD_DOA, "
    End If
    'hours
    SQLQ = SQLQ & "SUM(AD_HRS) AS AD_HRS, "
    'point
    SQLQ = SQLQ & "SUM(HR_ATTENDANCE.AD_POINT) AS AD_POINT,"
    SQLQ = SQLQ & Date_SQL(Date) & " AS AD_LDATE,"
    SQLQ = SQLQ & "'" & Time$ & "' AS AD_LTIME,"
    SQLQ = SQLQ & "'" & glbUserID & "' AS AD_LUSER, "
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HR_ATTENDANCE "
    SQLQ = SQLQ & "WHERE (AD_EMPNBR=" & EEID & " ) "
    SQLQ = SQLQ & "GROUP BY AD_REASON, AD_COMPNO,AD_EMPNBR,"
    If glbOracle Then
        SQLQ = SQLQ & " TO_CHAR(AD_DOA,'YYYY') "
    Else
        SQLQ = SQLQ & " Year(AD_DOA) "
    End If

Else
    xFList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE", "AD_ATT_ID")
    
    SQLQ = "INSERT INTO Term_ATTENDANCE (" & xFList & ",TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ", "
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HR_ATTENDANCE "
    SQLQ = SQLQ & "WHERE (AD_EMPNBR = " & EEID& & " )"
End If

gdbAdoIhr001.Execute SQLQ

TERM_ATTENDANCE = True

Exit Function

TERM_ATTENDANCE_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Attendance", "Term_Attendance", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_ATTENDANCE_HISTORY(EEID&, DtTm)
Dim SQLQ As String, dtATTD%, SQLQ2 As String
Dim iRow As Integer, Msg As String

TERM_ATTENDANCE_HISTORY = False
On Error GoTo TERM_ATTENDANCE_HISTORY_Err

'Ticket #27839 - Archived Attendance records Seniority Flag is gettting turned OFF when moved to TERM tables.
If glbchkSum = True Then  'laura nov 5, 1997
    SQLQ = "INSERT INTO Term_ATTENDANCE "
    SQLQ = SQLQ & "(AD_COMPNO,AD_EMPNBR,AD_DOA,AD_REASON,AD_HRS,AD_POINT,AD_LDATE,AD_LTIME,AD_LUSER,TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT  "
    SQLQ = SQLQ & "AH_COMPNO AS AD_COMPNO, "
    SQLQ = SQLQ & "AH_EMPNBR AS AD_EMPNBR, "
    If glbOracle Then
        SQLQ2 = " TO_DATE('31-12-' || TO_CHAR(AH_DOA,'YYYY'),'DD-MM-YYYY') AS AD_DOA, "
    ElseIf glbSQL Then
        SQLQ2 = " '" & IIf(glbFrench, "déc", "Dec") & " 31,'+CONVERT(varchar(4),AVG(Year(AH_DOA))) AS AD_DOA,"
    Else
        SQLQ2 = " CVDATE('" & IIf(glbFrench, "déc", "Dec") & " 31,'+STR(Year(AH_DOA))) As AD_DOA, "
    End If
    If glbchkSum = True Then  'laura dec 15, 1997
        SQLQ = SQLQ & SQLQ2
    Else
        SQLQ = SQLQ & "AH_DOA AS AD_DOA, "
    End If
    SQLQ = SQLQ & "AH_REASON AS AD_REASON, "
    If glbchkSum = True Then  'laura nov 5, 1997
        SQLQ = SQLQ & "SUM(AH_HRS) AS AD_HRS, "
    Else
        SQLQ = SQLQ & "AH_HRS AS AD_HRS, " 'laura nov 5, 1997
    End If
    'point
    SQLQ = SQLQ & "SUM(HR_ATTENDANCE_HISTORY.AH_POINT) AS AD_POINT,"    'Ticket #27839 - was missing here but was there in HR_Attendance to TERM
    SQLQ = SQLQ & Date_SQL(Date) & " AS AD_LDATE, "
    SQLQ = SQLQ & "'" & Time$ & "' AS AD_LTIME, "
    SQLQ = SQLQ & "'" & glbUserID & "' AS AD_LUSER, "
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HR_ATTENDANCE_HISTORY "
    SQLQ = SQLQ & "WHERE (AH_EMPNBR=" & EEID & " ) "
    If glbchkSum = True Then  'laura nov 5, 1997
        SQLQ = SQLQ & "GROUP BY AH_REASON, AH_COMPNO,AH_EMPNBR, "
        If glbOracle Then
            SQLQ = SQLQ & " TO_CHAR(AH_DOA,'YYYY') "
        Else
            SQLQ = SQLQ & " Year(AH_DOA) "
        End If
    End If
Else
    xFList = Get_Fields(gdbAdoIhr001, "HR_ATTENDANCE_HISTORY", "AH_ATT_ID")
    
    SQLQ = "INSERT INTO Term_ATTENDANCE (" & Replace(xFList, "AH_", "AD_") & ",TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ", "
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HR_ATTENDANCE_HISTORY "
    SQLQ = SQLQ & "WHERE (AH_EMPNBR = " & EEID& & " )"
End If

gdbAdoIhr001.Execute SQLQ

TERM_ATTENDANCE_HISTORY = True

Exit Function

TERM_ATTENDANCE_HISTORY_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Attendance_History", "Term_Attendance_History", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_BASIC(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_BASIC = False

On Error GoTo TERM_BASIC_Err
xFList = Get_Fields(gdbAdoIhr001, "HREMP", "")

SQLQ = "INSERT INTO Term_HREMP (" & xFList & ",TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ", "
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREMP "
SQLQ = SQLQ & "WHERE (HREMP.ED_EMPNBR = " & EEID& & " )"

gdbAdoIhr001.Execute SQLQ

TERM_BASIC = True

Exit Function

TERM_BASIC_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Basic", "Term_Basic", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_BENEFITS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_BENEFITS = False

On Error GoTo TERM_BENEFITS_Err
If glbWFC And IsDate(glbChgBenTermDate) Then
    xFList = Get_Fields(gdbAdoIhr001, "HRBENFT", "BF_BENE_ID,BF_CEASEDATE")
    SQLQ = "INSERT INTO Term_HRBENFT (" & xFList & ", TERM_SEQ,BF_CEASEDATE) "
Else
    xFList = Get_Fields(gdbAdoIhr001, "HRBENFT", "BF_BENE_ID")
    SQLQ = "INSERT INTO Term_HRBENFT (" & xFList & ", TERM_SEQ) "
End If
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
If glbWFC And IsDate(glbChgBenTermDate) Then
    SQLQ = SQLQ & "," & Date_SQL(glbChgBenTermDate) & " As BF_CEASEDATE "
End If
SQLQ = SQLQ & "FROM HRBENFT "
SQLQ = SQLQ & "WHERE (HRBENFT.BF_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ
xFList = Get_Fields(gdbAdoIhr001, "HRBENS", "BD_ID")
SQLQ = "INSERT INTO Term_HRBENS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRBENS "
SQLQ = SQLQ & "WHERE (HRBENS.BD_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_BENEFITS = True

Exit Function

TERM_BENEFITS_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Benefits", "Term_Benefits", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_COUNSEL(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_COUNSEL = False

On Error GoTo TERM_COUNSEL_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_COUNSEL", "CL_ID")
SQLQ = SQLQ & "INSERT INTO Term_HR_COUNSEL (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_COUNSEL "
SQLQ = SQLQ & "WHERE (HR_COUNSEL.CL_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_COUNSEL = True

Exit Function

TERM_COUNSEL_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_COUNSEL", "Term_COUNSEL", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HREEO(EEID As Long)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_HREEO = False

On Error GoTo TERM_EEO_Err

'check if there is a record in HREEO, delete the record with the same EO_EMPNBR
'Ticket #25969 Franks
SQLQ = "DELETE FROM Term_HREEO WHERE EO_EMPNBR = " & EEID& & " "
gdbAdoIhr001.Execute SQLQ
    
xFList = Get_Fields(gdbAdoIhr001, "HREEO", "") ' "EF_FOLLOWUP_ID")
SQLQ = SQLQ & "INSERT INTO Term_HREEO (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREEO "
SQLQ = SQLQ & "WHERE (HREEO.EO_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ


TERM_HREEO = True

Exit Function

TERM_EEO_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HREEO", "Term_HREEO", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_FOLLOW_UP(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_FOLLOW_UP = False

On Error GoTo TERM_FOLLOW_UP_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_FOLLOW_UP", "EF_FOLLOWUP_ID")
SQLQ = SQLQ & "INSERT INTO Term_FOLLOW_UP (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_FOLLOW_UP "
SQLQ = SQLQ & "WHERE (HR_FOLLOW_UP.EF_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_FOLLOW_UP = True

Exit Function

TERM_FOLLOW_UP_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_FOLLOW_UP", "Term_FOLLOW_UP", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HRDOC_EMP(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_EMP = False

On Error GoTo TERM_HRDOC_EMP_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EMP", "RE_ID") '"TERM_SEQ")
    SQLQ = "INSERT INTO TERM_HRDOC_EMP (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_EMP "
    SQLQ = SQLQ & "WHERE (HRDOC_EMP.RE_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='RESUME' AND RE_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    If Not rsDocActive.EOF Then
        SQLQ = "SELECT * FROM TERM_HRDOC_EMP WHERE RE_TYPE='RESUME' AND TERM_SEQ=" & glbTERM_Seq
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("RE_EMPNBR") = rsDocActive("RE_EMPNBR")
            rsDocTerm("RE_FILEEXT") = rsDocActive("RE_FILEEXT")
            rsDocTerm("RE_TYPE") = rsDocActive("RE_TYPE")
            rsDocTerm("RE_DOC") = rsDocActive("RE_DOC")
            rsDocTerm("RE_LDATE") = rsDocActive("RE_LDATE")
            rsDocTerm("RE_LTIME") = rsDocActive("RE_LTIME")
            rsDocTerm("RE_LUSER") = rsDocActive("RE_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
    End If
    rsDocActive.Close
    SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='TERMINATION' AND RE_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    If Not rsDocActive.EOF Then
        SQLQ = "SELECT * FROM TERM_HRDOC_EMP WHERE RE_TYPE='TERMINATION' AND TERM_SEQ=" & glbTERM_Seq
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("RE_EMPNBR") = rsDocActive("RE_EMPNBR")
            rsDocTerm("RE_FILEEXT") = rsDocActive("RE_FILEEXT")
            rsDocTerm("RE_TYPE") = rsDocActive("RE_TYPE")
            rsDocTerm("RE_DOC") = rsDocActive("RE_DOC")
            rsDocTerm("RE_LDATE") = rsDocActive("RE_LDATE")
            rsDocTerm("RE_LTIME") = rsDocActive("RE_LTIME")
            rsDocTerm("RE_LUSER") = rsDocActive("RE_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
    End If
    rsDocActive.Close
End If

SQLQ = "Delete FROM HRDOC_EMP "
SQLQ = SQLQ & " WHERE RE_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_EMP = True

Exit Function

TERM_HRDOC_EMP_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_EMP", "TERM_HRDOC_EMP", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_HREMP_OTHER(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_HRDOC_HREMP_OTHER = False

On Error GoTo TERM_HRDOC_HREMP_OTHER_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HREMP_OTHER", "ER_ID") '"TERM_SEQ")
    SQLQ = "INSERT INTO Term_HRDOC_HREMP_OTHER (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_HREMP_OTHER "
    SQLQ = SQLQ & "WHERE (HRDOC_HREMP_OTHER.ER_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_HREMP_OTHER WHERE ER_TYPE='OTHERINFO' AND ER_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    If Not rsDocActive.EOF Then
        SQLQ = "SELECT * FROM Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='OTHERINFO' AND TERM_SEQ=" & glbTERM_Seq
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("ER_EMPNBR") = rsDocActive("ER_EMPNBR")
            rsDocTerm("ER_FILEEXT") = rsDocActive("ER_FILEEXT")
            rsDocTerm("ER_TYPE") = rsDocActive("ER_TYPE")
            rsDocTerm("ER_DOC") = rsDocActive("ER_DOC")
            rsDocTerm("ER_LDATE") = rsDocActive("ER_LDATE")
            rsDocTerm("ER_LTIME") = rsDocActive("ER_LTIME")
            rsDocTerm("ER_LUSER") = rsDocActive("ER_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
    End If
    rsDocActive.Close
End If

SQLQ = "Delete FROM HRDOC_HREMP_OTHER "
SQLQ = SQLQ & " WHERE ER_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_HREMP_OTHER = True

Exit Function

TERM_HRDOC_HREMP_OTHER_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_HREMP_OTHER", "TERM_HRDOC_HREMP_OTHER", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_JOB_HISTORY(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_JOB_HISTORY = False

On Error GoTo TERM_HRDOC_JOB_HISTORY_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_JOB_HISTORY", "DJ_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_JOB_HISTORY (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_JOB_HISTORY "
    SQLQ = SQLQ & "WHERE (HRDOC_JOB_HISTORY.DJ_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE='OFFER' AND DJ_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_JOB_HISTORY WHERE DJ_TYPE='OFFER' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND DJ_JOB = '" & rsDocActive("DJ_JOB") & "' "
        If IsDate(rsDocActive("DJ_SDATE")) Then
            SQLQ = SQLQ & "AND DJ_SDATE = " & Date_SQL(rsDocActive("DJ_SDATE")) & " "
        End If
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DJ_EMPNBR") = rsDocActive("DJ_EMPNBR")
            rsDocTerm("DJ_SDATE") = rsDocActive("DJ_SDATE")
            rsDocTerm("DJ_JOB") = rsDocActive("DJ_JOB")
            rsDocTerm("DJ_DOC") = rsDocActive("DJ_DOC")
            rsDocTerm("DJ_FILEEXT") = rsDocActive("DJ_FILEEXT")
            rsDocTerm("DJ_TYPE") = rsDocActive("DJ_TYPE")
            rsDocTerm("DJ_LDATE") = rsDocActive("DJ_LDATE")
            rsDocTerm("DJ_LTIME") = rsDocActive("DJ_LTIME")
            rsDocTerm("DJ_LUSER") = rsDocActive("DJ_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_JOB_HISTORY "
SQLQ = SQLQ & "WHERE DJ_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_JOB_HISTORY = True

Exit Function

TERM_HRDOC_JOB_HISTORY_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_JOB_HISTORY", "TERM_HRDOC_JOB_HISTORY", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_COMMENTS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_COMMENTS = False

On Error GoTo TERM_HRDOC_COMMENTS_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_COMMENTS", "DO_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_COMMENTS (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_COMMENTS "
    SQLQ = SQLQ & "WHERE (HRDOC_COMMENTS.DO_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE='COMMENTS' AND DO_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_COMMENTS WHERE DO_TYPE='COMMENTS' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND DO_DOCKEY = " & rsDocActive("DO_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DO_EMPNBR") = rsDocActive("DO_EMPNBR")
            rsDocTerm("DO_COTYPE") = rsDocActive("DO_COTYPE")
            rsDocTerm("DO_EDATE") = rsDocActive("DO_EDATE")
            rsDocTerm("DO_DOC") = rsDocActive("DO_DOC")
            rsDocTerm("DO_FILEEXT") = rsDocActive("DO_FILEEXT")
            rsDocTerm("DO_TYPE") = rsDocActive("DO_TYPE")
            rsDocTerm("DO_LDATE") = rsDocActive("DO_LDATE")
            rsDocTerm("DO_LTIME") = rsDocActive("DO_LTIME")
            rsDocTerm("DO_LUSER") = rsDocActive("DO_LUSER")
            rsDocTerm("DO_DOCKEY") = rsDocActive("DO_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_COMMENTS "
SQLQ = SQLQ & "WHERE DO_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_COMMENTS = True

Exit Function

TERM_HRDOC_COMMENTS_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_COMMENTS", "TERM_HRDOC_COMMENTS", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_HEALTH_SAFETY_2(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_HEALTH_SAFETY_2 = False

On Error GoTo TERM_HRDOC_HEALTH_SAFETY_2_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HEALTH_SAFETY_2", "DE_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_HEALTH_SAFETY_2 (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_HEALTH_SAFETY_2 "
    SQLQ = SQLQ & "WHERE (HRDOC_HEALTH_SAFETY_2.DE_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='INCIDENT' AND DE_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='INCIDENT' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND DE_CASE = " & rsDocActive("DE_CASE") & " "
        SQLQ = SQLQ & "AND DE_DOCNO = " & rsDocActive("DE_DOCNO") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DE_EMPNBR") = rsDocActive("DE_EMPNBR")
            rsDocTerm("DE_CASE") = rsDocActive("DE_CASE")
            rsDocTerm("DE_OCCDATE") = rsDocActive("DE_OCCDATE")
            rsDocTerm("DE_DOCNO") = rsDocActive("DE_DOCNO")
            rsDocTerm("DE_DOCDESC") = rsDocActive("DE_DOCDESC")
            rsDocTerm("DE_DOC") = rsDocActive("DE_DOC")
            rsDocTerm("DE_FILEEXT") = rsDocActive("DE_FILEEXT")
            rsDocTerm("DE_TYPE") = rsDocActive("DE_TYPE")
            rsDocTerm("DE_LDATE") = rsDocActive("DE_LDATE")
            rsDocTerm("DE_LTIME") = rsDocActive("DE_LTIME")
            rsDocTerm("DE_LUSER") = rsDocActive("DE_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY_2 "
SQLQ = SQLQ & "WHERE DE_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_HEALTH_SAFETY_2 = True

Exit Function

TERM_HRDOC_HEALTH_SAFETY_2_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_HEALTH_SAFETY_2", "TERM_HRDOC_HEALTH_SAFETY_2", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_HEALTH_SAFETY(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_HEALTH_SAFETY = False

On Error GoTo TERM_HRDOC_HEALTH_SAFETY_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HEALTH_SAFETY", "DE_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_HEALTH_SAFETY (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_HEALTH_SAFETY "
    SQLQ = SQLQ & "WHERE (HRDOC_HEALTH_SAFETY.DE_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY WHERE DE_TYPE='INCIDENT' AND DE_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_HEALTH_SAFETY WHERE DE_TYPE='INCIDENT' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND DE_CASE = " & rsDocActive("DE_CASE") & " "
        SQLQ = SQLQ & "AND DE_DOCNO = " & rsDocActive("DE_DOCNO") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DE_EMPNBR") = rsDocActive("DE_EMPNBR")
            rsDocTerm("DE_CASE") = rsDocActive("DE_CASE")
            rsDocTerm("DE_OCCDATE") = rsDocActive("DE_OCCDATE")
            rsDocTerm("DE_DOCNO") = rsDocActive("DE_DOCNO")
            rsDocTerm("DE_DOCDESC") = rsDocActive("DE_DOCDESC")
            rsDocTerm("DE_DOC") = rsDocActive("DE_DOC")
            rsDocTerm("DE_FILEEXT") = rsDocActive("DE_FILEEXT")
            rsDocTerm("DE_TYPE") = rsDocActive("DE_TYPE")
            rsDocTerm("DE_LDATE") = rsDocActive("DE_LDATE")
            rsDocTerm("DE_LTIME") = rsDocActive("DE_LTIME")
            rsDocTerm("DE_LUSER") = rsDocActive("DE_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY "
SQLQ = SQLQ & "WHERE DE_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_HEALTH_SAFETY = True

Exit Function

TERM_HRDOC_HEALTH_SAFETY_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_HEALTH_SAFETY", "TERM_HRDOC_HEALTH_SAFETY", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7 = False

On Error GoTo TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HEALTH_SAFETY_CONCERNSWF7", "W7_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7 (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 "
    SQLQ = SQLQ & "WHERE (HRDOC_HEALTH_SAFETY_CONCERNSWF7.W7_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='INJURYWF7' AND W7_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='INJURYWF7' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND W7_CASE = " & rsDocActive("W7_CASE") & " "
        SQLQ = SQLQ & "AND W7_DOCKEY = " & rsDocActive("W7_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("W7_EMPNBR") = rsDocActive("W7_EMPNBR")
            rsDocTerm("W7_CASE") = rsDocActive("W7_CASE")
            rsDocTerm("W7_OCCDATE") = rsDocActive("W7_OCCDATE")
            rsDocTerm("W7_DOCKEY") = rsDocActive("W7_DOCKEY")
            rsDocTerm("W7_DOCDESC") = rsDocActive("W7_DOCDESC")
            rsDocTerm("W7_DOC") = rsDocActive("W7_DOC")
            rsDocTerm("W7_FILEEXT") = rsDocActive("W7_FILEEXT")
            rsDocTerm("W7_TYPE") = rsDocActive("W7_TYPE")
            rsDocTerm("W7_LDATE") = rsDocActive("W7_LDATE")
            rsDocTerm("W7_LTIME") = rsDocActive("W7_LTIME")
            rsDocTerm("W7_LUSER") = rsDocActive("W7_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 "
SQLQ = SQLQ & "WHERE W7_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7 = True

Exit Function

TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7", "TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_OHS_WRITTEN_OFFER(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_HRDOC_OHS_WRITTEN_OFFER = False

On Error GoTo TERM_HRDOC_OHS_WRITTEN_OFFER_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_OHS_WRITTEN_OFFER", "F7_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_OHS_WRITTEN_OFFER (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_OHS_WRITTEN_OFFER "
    SQLQ = SQLQ & "WHERE (HRDOC_OHS_WRITTEN_OFFER.F7_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='INJURYWF7_WRITTENOFR' AND F7_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='INJURYWF7_WRITTENOFR' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND F7_CASE = " & rsDocActive("F7_CASE") & " "
        SQLQ = SQLQ & "AND F7_DOCKEY = " & rsDocActive("F7_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("F7_EMPNBR") = rsDocActive("F7_EMPNBR")
            rsDocTerm("F7_CASE") = rsDocActive("F7_CASE")
            rsDocTerm("F7_OCCDATE") = rsDocActive("F7_OCCDATE")
            rsDocTerm("F7_DOCKEY") = rsDocActive("F7_DOCKEY")
            rsDocTerm("F7_DOCDESC") = rsDocActive("F7_DOCDESC")
            rsDocTerm("F7_DOC") = rsDocActive("F7_DOC")
            rsDocTerm("F7_FILEEXT") = rsDocActive("F7_FILEEXT")
            rsDocTerm("F7_TYPE") = rsDocActive("F7_TYPE")
            rsDocTerm("F7_LDATE") = rsDocActive("F7_LDATE")
            rsDocTerm("F7_LTIME") = rsDocActive("F7_LTIME")
            rsDocTerm("F7_LUSER") = rsDocActive("F7_LUSER")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_OHS_WRITTEN_OFFER "
SQLQ = SQLQ & "WHERE F7_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_OHS_WRITTEN_OFFER = True

Exit Function

TERM_HRDOC_OHS_WRITTEN_OFFER_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_OHS_WRITTEN_OFFER", "TERM_HRDOC_OHS_WRITTEN_OFFER", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_HRDOLENT(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_HRDOLENT = False

On Error GoTo TERM_HRDOC_HRDOLENT_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HRDOLENT", "DE_ID")
    SQLQ = "INSERT INTO Term_HRDOC_DOLENT (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_HRDOLENT "
    SQLQ = SQLQ & "WHERE (HRDOC_HRDOLENT.DE_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_HRDOLENT WHERE DE_TYPE='DOLLARENT' AND DE_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM Term_HRDOC_DOLENT WHERE DE_TYPE='DOLLARENT' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND DE_DOCKEY = " & rsDocActive("DE_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DE_EMPNBR") = rsDocActive("DE_EMPNBR")
            rsDocTerm("DE_CLTYPE") = rsDocActive("DE_CLTYPE")
            rsDocTerm("DE_COUDATE") = rsDocActive("DE_COUDATE")
            rsDocTerm("DE_DOC") = rsDocActive("DE_DOC")
            rsDocTerm("DE_FILEEXT") = rsDocActive("DE_FILEEXT")
            rsDocTerm("DE_TYPE") = rsDocActive("DE_TYPE")
            rsDocTerm("DE_LDATE") = rsDocActive("DE_LDATE")
            rsDocTerm("DE_LTIME") = rsDocActive("DE_LTIME")
            rsDocTerm("DE_LUSER") = rsDocActive("DE_LUSER")
            rsDocTerm("DE_DOCKEY") = rsDocActive("DE_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_HRDOLENT "
SQLQ = SQLQ & "WHERE DE_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_HRDOLENT = True

Exit Function

TERM_HRDOC_HRDOLENT_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HRDOC_DOLENT", "Term_HRDOC_DOLENT", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HRDOC_HREDU(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_HREDU = False

On Error GoTo TERM_HRDOC_HREDU_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HREDU", "EU_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_HREDU (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_HREDU "
    SQLQ = SQLQ & "WHERE (HRDOC_HREDU.EU_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_HREDU WHERE EU_TYPE='EDSEM' AND EU_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_HREDU WHERE EU_TYPE='EDSEM' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND EU_DOCKEY = " & rsDocActive("EU_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("EU_EMPNBR") = rsDocActive("EU_EMPNBR")
            rsDocTerm("EU_CLTYPE") = rsDocActive("EU_CLTYPE")
            rsDocTerm("EU_COUDATE") = rsDocActive("EU_COUDATE")
            rsDocTerm("EU_DOC") = rsDocActive("EU_DOC")
            rsDocTerm("EU_FILEEXT") = rsDocActive("EU_FILEEXT")
            rsDocTerm("EU_TYPE") = rsDocActive("EU_TYPE")
            rsDocTerm("EU_LDATE") = rsDocActive("EU_LDATE")
            rsDocTerm("EU_LTIME") = rsDocActive("EU_LTIME")
            rsDocTerm("EU_LUSER") = rsDocActive("EU_LUSER")
            rsDocTerm("EU_DOCKEY") = rsDocActive("EU_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_HREDU "
SQLQ = SQLQ & "WHERE EU_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_HREDU = True

Exit Function

TERM_HRDOC_HREDU_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_HREDU", "TERM_HRDOC_HREDU", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HRDOC_EDSEM(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_EDSEM = False

On Error GoTo TERM_HRDOC_EDSEM_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EDSEM", "ES_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_EDSEM (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_EDSEM "
    SQLQ = SQLQ & "WHERE (HRDOC_EDSEM.ES_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_EDSEM WHERE ES_TYPE='EDSEM' AND ES_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_EDSEM WHERE ES_TYPE='EDSEM' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND ES_DOCKEY = " & rsDocActive("ES_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("ES_EMPNBR") = rsDocActive("ES_EMPNBR")
            'Ticket #15131, there were not these two fields in databases
            'rsDocTerm("ES_CLTYPE") = rsDocActive("ES_CLTYPE")
            'rsDocTerm("ES_COUDATE") = rsDocActive("ES_COUDATE")
            rsDocTerm("ES_DOC") = rsDocActive("ES_DOC")
            rsDocTerm("ES_FILEEXT") = rsDocActive("ES_FILEEXT")
            rsDocTerm("ES_TYPE") = rsDocActive("ES_TYPE")
            rsDocTerm("ES_LDATE") = rsDocActive("ES_LDATE")
            rsDocTerm("ES_LTIME") = rsDocActive("ES_LTIME")
            rsDocTerm("ES_LUSER") = rsDocActive("ES_LUSER")
            rsDocTerm("ES_DOCKEY") = rsDocActive("ES_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_EDSEM "
SQLQ = SQLQ & "WHERE ES_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_EDSEM = True

Exit Function

TERM_HRDOC_EDSEM_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_EDSEM", "TERM_HRDOC_EDSEM", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_EDSEM_RETEST(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_EDSEM_RETEST = False

On Error GoTo TERM_HRDOC_EDSEM_RETEST_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EDSEM_RETEST", "ES_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_EDSEM_RETEST (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_EDSEM_RETEST "
    SQLQ = SQLQ & "WHERE (HRDOC_EDSEM_RETEST.ES_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_EDSEM_RETEST WHERE ES_TYPE='EDSEM' AND ES_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_EDSEM_RETEST WHERE ES_TYPE='EDSEM' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND ES_DOCKEY = " & rsDocActive("ES_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("ES_EMPNBR") = rsDocActive("ES_EMPNBR")
            rsDocTerm("ES_CLTYPE") = rsDocActive("ES_CLTYPE")
            rsDocTerm("ES_COUDATE") = rsDocActive("ES_COUDATE")
            rsDocTerm("ES_DOC") = rsDocActive("ES_DOC")
            rsDocTerm("ES_FILEEXT") = rsDocActive("ES_FILEEXT")
            rsDocTerm("ES_TYPE") = rsDocActive("ES_TYPE")
            rsDocTerm("ES_LDATE") = rsDocActive("ES_LDATE")
            rsDocTerm("ES_LTIME") = rsDocActive("ES_LTIME")
            rsDocTerm("ES_LUSER") = rsDocActive("ES_LUSER")
            rsDocTerm("ES_DOCKEY") = rsDocActive("ES_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_EDSEM_RETEST "
SQLQ = SQLQ & "WHERE ES_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_EDSEM_RETEST = True

Exit Function

TERM_HRDOC_EDSEM_RETEST_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_EDSEM_RETEST", "TERM_HRDOC_EDSEM_RETEST", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_COUNSEL(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_COUNSEL = False

On Error GoTo TERM_HRDOC_COUNSEL_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_COUNSEL", "DC_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_COUNSEL (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_COUNSEL "
    SQLQ = SQLQ & "WHERE (HRDOC_COUNSEL.DC_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE='COUNSEL' AND DC_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_COUNSEL WHERE DC_TYPE='COUNSEL' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND DC_DOCKEY = " & rsDocActive("DC_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DC_EMPNBR") = rsDocActive("DC_EMPNBR")
            rsDocTerm("DC_CLTYPE") = rsDocActive("DC_CLTYPE")
            rsDocTerm("DC_COUDATE") = rsDocActive("DC_COUDATE")
            rsDocTerm("DC_DOC") = rsDocActive("DC_DOC")
            rsDocTerm("DC_FILEEXT") = rsDocActive("DC_FILEEXT")
            rsDocTerm("DC_TYPE") = rsDocActive("DC_TYPE")
            rsDocTerm("DC_LDATE") = rsDocActive("DC_LDATE")
            rsDocTerm("DC_LTIME") = rsDocActive("DC_LTIME")
            rsDocTerm("DC_LUSER") = rsDocActive("DC_LUSER")
            rsDocTerm("DC_DOCKEY") = rsDocActive("DC_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_COUNSEL "
SQLQ = SQLQ & "WHERE DC_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_COUNSEL = True

Exit Function

TERM_HRDOC_COUNSEL_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_COUNSEL", "TERM_HRDOC_COUNSEL", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_PERFORM_HISTORY(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HRDOC_PERFORM_HISTORY = False

On Error GoTo TERM_HRDOC_PERFORM_HISTORY_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_PERFORM_HISTORY", "DH_ID")
    SQLQ = "INSERT INTO TERM_HRDOC_PERFORM_HISTORY (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_PERFORM_HISTORY "
    SQLQ = SQLQ & "WHERE (HRDOC_PERFORM_HISTORY.DH_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE='PERFORMANCE' AND DH_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM TERM_HRDOC_PERFORM_HISTORY WHERE DH_TYPE='PERFORMANCE' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND DH_DOCKEY = " & rsDocActive("DH_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DH_EMPNBR") = rsDocActive("DH_EMPNBR")
            rsDocTerm("DH_JOB") = rsDocActive("DH_JOB")
            rsDocTerm("DH_PREVDATE") = rsDocActive("DH_PREVDATE")
            rsDocTerm("DH_DOC") = rsDocActive("DH_DOC")
            rsDocTerm("DH_FILEEXT") = rsDocActive("DH_FILEEXT")
            rsDocTerm("DH_TYPE") = rsDocActive("DH_TYPE")
            rsDocTerm("DH_LDATE") = rsDocActive("DH_LDATE")
            rsDocTerm("DH_LTIME") = rsDocActive("DH_LTIME")
            rsDocTerm("DH_LUSER") = rsDocActive("DH_LUSER")
            rsDocTerm("DH_DOCKEY") = rsDocActive("DH_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If


SQLQ = "DELETE FROM HRDOC_PERFORM_HISTORY "
SQLQ = SQLQ & "WHERE DH_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_PERFORM_HISTORY = True

Exit Function

TERM_HRDOC_PERFORM_HISTORY_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDOC_PERFORM_HISTORY", "TERM_HRDOC_PERFORM_HISTORY", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRDOC_TRADE(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_HRDOC_TRADE = False

On Error GoTo TERM_HRDOC_TRADE_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_TRADE", "TD_ID")
    SQLQ = "INSERT INTO Term_HRDOC_TRADE (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_TRADE "
    SQLQ = SQLQ & "WHERE (HRDOC_TRADE.TD_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_TRADE WHERE TD_TYPE='ASSOCIATIONS' AND TD_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM Term_HRDOC_TRADE WHERE TD_TYPE='ASSOCIATIONS' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND TD_DOCKEY = " & rsDocActive("TD_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("TD_EMPNBR") = rsDocActive("TD_EMPNBR")
            rsDocTerm("TD_CODE") = rsDocActive("TD_CODE")
            rsDocTerm("TD_BEGINDT") = rsDocActive("TD_BEGINDT")
            rsDocTerm("TD_DOC") = rsDocActive("TD_DOC")
            rsDocTerm("TD_FILEEXT") = rsDocActive("TD_FILEEXT")
            rsDocTerm("TD_TYPE") = rsDocActive("TD_TYPE")
            rsDocTerm("TD_LDATE") = rsDocActive("TD_LDATE")
            rsDocTerm("TD_LTIME") = rsDocActive("TD_LTIME")
            rsDocTerm("TD_LUSER") = rsDocActive("TD_LUSER")
            rsDocTerm("TD_DOCKEY") = rsDocActive("TD_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_TRADE "
SQLQ = SQLQ & "WHERE TD_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_TRADE = True

Exit Function

TERM_HRDOC_TRADE_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HRDOC_TRADE", "Term_HRDOC_TRADE", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HRDOC_ATTENDANCE(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_HRDOC_ATTENDANCE = False

On Error GoTo TERM_HRDOC_ATTENDANCE_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_ATTENDANCE", "AD_ID")
    SQLQ = "INSERT INTO Term_HRDOC_ATTENDANCE (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_ATTENDANCE "
    SQLQ = SQLQ & "WHERE (HRDOC_ATTENDANCE.AD_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_TYPE='ATTENDANCE' AND AD_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM Term_HRDOC_ATTENDANCE WHERE AD_TYPE='ATTENDANCE' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND AD_DOCKEY = " & rsDocActive("AD_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("AD_EMPNBR") = rsDocActive("AD_EMPNBR")
            rsDocTerm("AD_REASON") = rsDocActive("AD_REASON")
            rsDocTerm("AD_DOA") = rsDocActive("AD_DOA")
            rsDocTerm("AD_DOC") = rsDocActive("AD_DOC")
            rsDocTerm("AD_FILEEXT") = rsDocActive("AD_FILEEXT")
            rsDocTerm("AD_TYPE") = rsDocActive("AD_TYPE")
            rsDocTerm("AD_LDATE") = rsDocActive("AD_LDATE")
            rsDocTerm("AD_LTIME") = rsDocActive("AD_LTIME")
            rsDocTerm("AD_LUSER") = rsDocActive("AD_LUSER")
            rsDocTerm("AD_DOCKEY") = rsDocActive("AD_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_ATTENDANCE "
SQLQ = SQLQ & "WHERE AD_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_ATTENDANCE = True

Exit Function

TERM_HRDOC_ATTENDANCE_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HRDOC_ATTENDANCE", "Term_HRDOC_ATTENDANCE", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HRDOC_EMP_FLAGS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_HRDOC_EMP_FLAGS = False

On Error GoTo TERM_HRDOC_EMP_FLAGS_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EMP_FLAGS", "EF_ID")
    SQLQ = "INSERT INTO Term_HRDOC_EMP_FLAGS (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HRDOC_EMP_FLAGS "
    SQLQ = SQLQ & "WHERE (HRDOC_EMP_FLAGS.EF_EMPNBR=" & EEID & " )"
    gdbAdoIhr001_DOC.Execute SQLQ
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_TYPE='EMPLOYEEFLAG' AND EF_EMPNBR=" & EEID
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM Term_HRDOC_EMP_FLAGS WHERE EF_TYPE='EMPLOYEEFLAG' AND TERM_SEQ=" & glbTERM_Seq & " "
        SQLQ = SQLQ & "AND EF_DOCKEY = " & rsDocActive("EF_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("EF_EMPNBR") = rsDocActive("EF_EMPNBR")
            rsDocTerm("EF_FLAG") = rsDocActive("EF_FLAG")
            rsDocTerm("EF_FLAGDTE") = rsDocActive("EF_FLAGDTE")
            rsDocTerm("EF_DOC") = rsDocActive("EF_DOC")
            rsDocTerm("EF_FILEEXT") = rsDocActive("EF_FILEEXT")
            rsDocTerm("EF_TYPE") = rsDocActive("EF_TYPE")
            rsDocTerm("EF_LDATE") = rsDocActive("EF_LDATE")
            rsDocTerm("EF_LTIME") = rsDocActive("EF_LTIME")
            rsDocTerm("EF_LUSER") = rsDocActive("EF_LUSER")
            rsDocTerm("EF_DOCKEY") = rsDocActive("EF_DOCKEY")
            rsDocTerm("TERM_SEQ") = glbTERM_Seq
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM HRDOC_EMP_FLAGS "
SQLQ = SQLQ & "WHERE EF_EMPNBR=" & EEID

gdbAdoIhr001_DOC.Execute SQLQ

TERM_HRDOC_EMP_FLAGS = True

Exit Function

TERM_HRDOC_EMP_FLAGS_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HRDOC_EMP_FLAGS", "Term_HRDOC_EMP_FLAGS", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_EMPADP(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EMPADP = False

On Error GoTo TERM_EMPADP_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_ADP", "AP_ID")
SQLQ = "INSERT INTO Term_HR_ADP (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_ADP "
SQLQ = SQLQ & "WHERE (HR_ADP.AP_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_EMPADP = True

Exit Function

TERM_EMPADP_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HR_ADP", "Term_HR_ADP", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_PROFIT_SHARING(EEID As Long) 'Ticket #20052
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_PROFIT_SHARING = False

On Error GoTo TERM_EMPADP_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_PROFIT_SHARING", "PS_ID")
SQLQ = "INSERT INTO Term_PROFIT_SHARING (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_PROFIT_SHARING "
SQLQ = SQLQ & "WHERE (HR_PROFIT_SHARING.PS_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_PROFIT_SHARING = True

Exit Function

TERM_EMPADP_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_PROFIT_SHARING", "Term_PROFIT_SHARING", "Insert - " & SQLQ)
'Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
    Resume Next
'End If


End Function
Function TERM_EMPPAYROLL_TRANSACTION(EEID As Long) 'Ticket #18232
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EMPPAYROLL_TRANSACTION = False

On Error GoTo TERM_EMPADP_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_PAYROLL_TRANSACTION", "PT_ID")
SQLQ = "INSERT INTO Term_PAYROLL_TRANSACTION (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_PAYROLL_TRANSACTION "
SQLQ = SQLQ & "WHERE (HR_PAYROLL_TRANSACTION.PT_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_EMPPAYROLL_TRANSACTION = True

Exit Function

TERM_EMPADP_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_PAYROLL_TRANSACTION", "Term_PAYROLL_TRANSACTION", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HRP_PENSION_MEMBERSHIP(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

On Error GoTo TERM_PenTable_Err

TERM_HRP_PENSION_MEMBERSHIP = False

xFList = Get_Fields(gdbAdoIhr001, "HRP_PENSION_MEMBERSHIP", "PE_ID,TERM_SEQ")
SQLQ = "INSERT INTO Term_HRP_PENSION_MEMBERSHIP (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRP_PENSION_MEMBERSHIP "
SQLQ = SQLQ & "WHERE (HRP_PENSION_MEMBERSHIP.PE_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_HRP_PENSION_MEMBERSHIP = True

Exit Function

TERM_PenTable_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRP_PENSION_MEMBERSHIP", "TERM_HRP_PENSION_MEMBERSHIP", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function
Function TERM_HRP_PENSION_MASTER(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

On Error GoTo TERM_PenTable_Err

TERM_HRP_PENSION_MASTER = False

xFList = Get_Fields(gdbAdoIhr001, "HRP_PENSION_MASTER", "PE_ID,TERM_SEQ")
SQLQ = "INSERT INTO Term_HRP_PENSION_MASTER (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRP_PENSION_MASTER "
SQLQ = SQLQ & "WHERE (HRP_PENSION_MASTER.PE_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_HRP_PENSION_MASTER = True

Exit Function

TERM_PenTable_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRP_PENSION_MASTER", "TERM_HRP_PENSION_MASTER", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function
Function TERM_HRP_PENSION_BENEFICIARY(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

On Error GoTo TERM_PenTable_Err

TERM_HRP_PENSION_BENEFICIARY = False

xFList = Get_Fields(gdbAdoIhr001, "HRP_PENSION_BENEFICIARY", "PE_ID,TERM_SEQ")
SQLQ = "INSERT INTO Term_HRP_PENSION_BENEFICIARY (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRP_PENSION_BENEFICIARY "
SQLQ = SQLQ & "WHERE (HRP_PENSION_BENEFICIARY.PE_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_HRP_PENSION_BENEFICIARY = True

Exit Function

TERM_PenTable_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRP_PENSION_BENEFICIARY", "TERM_HRP_PENSION_BENEFICIARY", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HRP_PA_MASTER(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

On Error GoTo TERM_PenTable_Err

TERM_HRP_PA_MASTER = False

xFList = Get_Fields(gdbAdoIhr001, "HRP_PA_MASTER", "PE_ID,TERM_SEQ")
SQLQ = "INSERT INTO TERM_HRP_PA_MASTER (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRP_PA_MASTER "
SQLQ = SQLQ & "WHERE (HRP_PA_MASTER.PE_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_HRP_PA_MASTER = True

Exit Function

TERM_PenTable_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRP_PA_MASTER", "TERM_HRP_PA_MASTER", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function
Function TERM_HRP_PA_DETAILS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

On Error GoTo TERM_PenTable_Err

TERM_HRP_PA_DETAILS = False

xFList = Get_Fields(gdbAdoIhr001, "HRP_PA_DETAILS", "PE_ID,TERM_SEQ")
SQLQ = "INSERT INTO Term_HRP_PA_DETAILS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRP_PA_DETAILS "
SQLQ = SQLQ & "WHERE (HRP_PA_DETAILS.PE_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_HRP_PA_DETAILS = True

Exit Function

TERM_PenTable_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRP_PA_DETAILS", "TERM_HRP_PA_DETAILS", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_EMPOTHER(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EMPOTHER = False

On Error GoTo TERM_EMPOTHER_Err

xFList = Get_Fields(gdbAdoIhr001, "HREMP_OTHER", "TERM_SEQ")
SQLQ = "INSERT INTO Term_HREMP_OTHER (" & xFList & ",TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREMP_OTHER "
SQLQ = SQLQ & "WHERE (HREMP_OTHER.ER_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_EMPOTHER = True

Exit Function

TERM_EMPOTHER_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HREMP_OTHER", "Term_HREMP_OTHER", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_HREMPHIS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HREMPHIS = False

On Error GoTo TERM_HREMPHIS_Err

xFList = Get_Fields(gdbAdoIhr001, "HREMPHIS", "EE_ID,TERM_SEQ")
SQLQ = "INSERT INTO Term_HREMPHIS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREMPHIS "
SQLQ = SQLQ & "WHERE (HREMPHIS.EE_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_HREMPHIS = True

Exit Function

TERM_HREMPHIS_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HREMPHIS", "Term_HREMPHIS", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_USERDEFINE_TABLE(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_USERDEFINE_TABLE = False

On Error GoTo TERM_USERDEFINE_TABLE_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_USERDEFINE_TABLE", "UD_ID")
SQLQ = "INSERT INTO Term_USERDEFINE_TABLE (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_USERDEFINE_TABLE "
SQLQ = SQLQ & "WHERE (HR_USERDEFINE_TABLE.UD_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_USERDEFINE_TABLE = True

Exit Function

TERM_USERDEFINE_TABLE_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_USERDEFINE_TABLE", "Term_USERDEFINE_TABLE", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_EDUCSEM(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EDUCSEM = False

On Error GoTo TERM_EDUCSEM_Err
xFList = Get_Fields(gdbAdoIhr001, "HREDSEM", "ES_ID")
SQLQ = "INSERT INTO Term_HREDSEM (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREDSEM "
SQLQ = SQLQ & "WHERE (HREDSEM.ES_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


TERM_EDUCSEM = True

Exit Function

TERM_EDUCSEM_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HREDSEM", "Term_HREDSEM", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function
Function TERM_HealthReoccurrence(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HealthReoccurrence = False

On Error GoTo TERM_HealthCost_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_OHS_REOCCURENCE", "CC_WCBC_ID")
SQLQ = SQLQ & "INSERT INTO Term_OHS_REOCCURENCE (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OHS_REOCCURENCE "
SQLQ = SQLQ & "WHERE (HR_OHS_REOCCURENCE.CC_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_HealthReoccurrence = True

Exit Function

TERM_HealthCost_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HealthCost", "Term_HealthCost", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function
Function TERM_HealthCost(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HealthCost = False

On Error GoTo TERM_HealthCost_Err
xFList = Get_Fields(gdbAdoIhr001, "HROHSCOS", "CC_WCBC_ID")
SQLQ = SQLQ & "INSERT INTO Term_HROHSCOS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HROHSCOS "
SQLQ = SQLQ & "WHERE (HROHSCOS.CC_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_HealthCost = True

Exit Function

TERM_HealthCost_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HealthCost", "Term_HealthCost", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_HealthSafety(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_HealthSafety = False

On Error GoTo TERM_HealthSafety_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_OCC_HEALTH_SAFETY", "EC_ID")
SQLQ = SQLQ & "INSERT INTO Term_HR_OCC_HEALTH_SAFETY (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OCC_HEALTH_SAFETY "
SQLQ = SQLQ & "WHERE (HR_OCC_HEALTH_SAFETY.EC_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_HealthSafety = True

Exit Function

TERM_HealthSafety_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HealthSafety", "Term_HealthSafety", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_JOB(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_JOB = False
On Error GoTo TERM_JOB_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_JOB_HISTORY", "JH_ID")

SQLQ = "INSERT INTO Term_JOB_HISTORY (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ", " & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_JOB_HISTORY "
SQLQ = SQLQ & "WHERE (HR_JOB_HISTORY.JH_EMPNBR= " & EEID & " )"
gdbAdoIhr001.Execute SQLQ

'Friesens - Ticket #16189
If glbCompSerial = "S/N - 2279W" Then
    xFList = Get_Fields(gdbAdoIhr001, "HR_TEMP_WORK", "TW_ID")
    SQLQ = ""
    SQLQ = SQLQ & "INSERT INTO Term_TEMP_WORK (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT  " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HR_TEMP_WORK "
    SQLQ = SQLQ & "WHERE (HR_TEMP_WORK.TW_EMPNBR=" & EEID & " )"
    
    gdbAdoIhr001.Execute SQLQ
End If


TERM_JOB = True

Exit Function

TERM_JOB_Err:
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Job", "Term_Job", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Function

Function TERM_LIST(EEID&, DtTm, TReason$, TRDesc$, TComment$, TRehire$, Optional TCause)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim strTm$

TERM_LIST = False

On Error GoTo TERM_LIST_Err
strTm$ = Time$
SQLQ = "INSERT INTO Term_HRTRMEMP "
SQLQ = SQLQ & "(Company, Employee_Number, "
SQLQ = SQLQ & "Term_DOT, Term_Reason, Term_Rehire, Term_Comments, "
If Not IsMissing(TCause) Then
    If Len(TCause) > 0 Then
        SQLQ = SQLQ & "Term_Cause,"
    End If
End If
SQLQ = SQLQ & "Term_LDATE, Term_LTIME, Term_LUSER,TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)


SQLQ = SQLQ & "SELECT   "
SQLQ = SQLQ & "ED_COMPNO AS Company, ED_EMPNBR AS Employee_Number, "
SQLQ = SQLQ & Date_SQL(DtTm) & " AS Term_DOT, "
SQLQ = SQLQ & "'" & TReason & "' AS Term_Reason, "
SQLQ = SQLQ & "'" & IIf(TRehire$ = "Yes", "1", "0") & "' AS Term_Rehire, "
SQLQ = SQLQ & "'" & Replace(TComment, "'", "''") & "' AS Term_Comments, "
If Not IsMissing(TCause) Then
    If Len(TCause) > 0 Then
        SQLQ = SQLQ & "'" & TCause & "' AS Term_Cause, "
    End If
End If
SQLQ = SQLQ & Date_SQL(Date) & " As Term_LDATE, "
SQLQ = SQLQ & "'" & strTm & "' as Term_LTIME, "
SQLQ = SQLQ & "'" & glbUserID & "' as Term_LUSER, "
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREMP "
SQLQ = SQLQ & "WHERE HREMP.ED_EMPNBR = " & EEID
gdbAdoIhr001.Execute SQLQ

TERM_LIST = True

Exit Function

TERM_LIST_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_List", "Term_HRTRMEMP", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_PERFORM(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_PERFORM = False
On Error GoTo TERM_PERFORM_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_PERFORM_HISTORY", "PH_ID")
SQLQ = SQLQ & "INSERT INTO Term_PERFORM_HISTORY (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_PERFORM_HISTORY "
SQLQ = SQLQ & "WHERE (HR_PERFORM_HISTORY.PH_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ


If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
    xFList = Get_Fields(gdbAdoIhr001, "HR_PERFORM_FRIESEN", "PH_ID")
    SQLQ = ""
    SQLQ = SQLQ & "INSERT INTO Term_PERFORM_FRIESEN (" & xFList & ", TERM_SEQ) "
    SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
    SQLQ = SQLQ & "SELECT  " & xFList & ","
    SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
    SQLQ = SQLQ & "FROM HR_PERFORM_FRIESEN "
    SQLQ = SQLQ & "WHERE (HR_PERFORM_FRIESEN.PH_EMPNBR=" & EEID & " )"
    
    gdbAdoIhr001.Execute SQLQ
End If


TERM_PERFORM = True

Exit Function

TERM_PERFORM_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Perform", "Term_Perform", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Function

Function TERM_SALARY(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_SALARY = False

On Error GoTo TERM_SALARY_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_SALARY_HISTORY", "SH_ID")
SQLQ = SQLQ & "INSERT INTO Term_SALARY_HISTORY (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_SALARY_HISTORY "
SQLQ = SQLQ & "WHERE (HR_SALARY_HISTORY.SH_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_SALARY = True

Exit Function

TERM_SALARY_Err:

glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Salary", "Term_Salary", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Function


Function TERM_OHS_Contact(EEID As Long)      'JDY 1/25/99
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_OHS_Contact = False

On Error GoTo TERM_OHS_Contact_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_OHS_CONTACT", "CT_ID")
SQLQ = SQLQ & "INSERT INTO Term_OHS_CONTACT (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OHS_CONTACT "
SQLQ = SQLQ & "WHERE (HR_OHS_CONTACT.CT_Empnbr=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_OHS_Contact = True

Exit Function

TERM_OHS_Contact_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_OHS_Contact", "Term_OHS_Contact", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function



Function REIN_ATTENDANCE(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_ATTENDANCE = False

On Error GoTo REIN_ATTENDANCE_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_ATTENDANCE", "TERM_SEQ,AD_ATT_ID,AD_EMPNBR")

SQLQ = "INSERT INTO HR_ATTENDANCE (" & xFList & ", AD_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS AD_EMPNBR "
SQLQ = SQLQ & " FROM Term_ATTENDANCE WHERE (Term_ATTENDANCE.TERM_SEQ= " & EESEQ & ")"


gdbAdoIhr001X.Execute SQLQ
REIN_ATTENDANCE = True


Exit Function

REIN_ATTENDANCE_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_ATTENDANCE", "REIN_ATTENDANCE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_BASIC(EEID As Long, EESEQ As Long, TermDate As String, Optional xPayID)
Dim SQLQ As String, mQry As QueryDef
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_BASIC = False

On Error GoTo REIN_BASIC_Err
If glbCompSerial = "S/N - 2192W" Or glbWFC Or glbCompSerial = "S/N - 2370W" Then     'County of Essex
    xFList = Get_Fields(gdbAdoIhr001X, "Term_HREMP", "TERM_SEQ,ED_ID,ED_EMPNBR,ED_PAYROLL_ID")
    
    SQLQ = "INSERT INTO HREMP (" & xFList & ",ED_EMPNBR,ED_PAYROLL_ID) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS ED_EMPNBR "
    SQLQ = SQLQ & ", '" & xPayID & "' AS ED_PAYROLL_ID "
    SQLQ = SQLQ & " FROM Term_HREMP WHERE (Term_HREMP.TERM_SEQ= " & EESEQ & ")"
Else
    xFList = Get_Fields(gdbAdoIhr001X, "Term_HREMP", "TERM_SEQ,ED_ID,ED_EMPNBR")
    
    SQLQ = "INSERT INTO HREMP (" & xFList & ",ED_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS ED_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HREMP WHERE (Term_HREMP.TERM_SEQ= " & EESEQ & ")"
End If
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

REIN_BASIC = True

'gdbAdoIhr001X.Close
Exit Function
REIN_BASIC_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_BASIC", "REIN_BASIC", "Insert")
Screen.MousePointer = DEFAULT
Exit Function
'If gintRollBack% = False Then
'    Resume Next
'End If

End Function

Function REIN_BENEFITS(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
Dim rs As New ADODB.Recordset

REIN_BENEFITS = False

On Error GoTo REIN_BENEFITS_Err
If glbVadim Then
    gdbAdoIhr001X.BeginTrans
    gdbAdoIhr001X.Execute "Update Term_HRBENFT SET BF_LUSER='VADIM_INTEGRATION' WHERE Term_HRBENFT.TERM_SEQ= " & EESEQ
    gdbAdoIhr001X.CommitTrans
End If

xFList = Get_Fields(gdbAdoIhr001X, "Term_HRBENFT", "TERM_SEQ,BF_BENE_ID,BF_EMPNBR")

SQLQ = "INSERT INTO HRBENFT (" & xFList & ", BF_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS BF_EMPNBR "
SQLQ = SQLQ & " FROM Term_HRBENFT WHERE (Term_HRBENFT.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

If glbLinamar Then 'added by Bryan 17/Apr/2006 Ticket#10688
    SQLQ = "SELECT BF_BCODE FROM HRBENFT WHERE BF_EMPNBR=" & CStr(EEID)
    rs.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
    If rs.EOF = False Then
        Do
            rs("BF_BCODE") = CStr(Right(EEID, 3)) & Right(rs("BF_BCODE"), Len(rs("BF_BCODE")) - 3)
            rs.Update
            rs.MoveNext
        Loop Until rs.EOF
    End If
    rs.Close
End If
If glbVadim Then
    gdbAdoIhr001X.BeginTrans
    gdbAdoIhr001X.Execute "Update HRBENFT SET BF_LUSER='" & glbUserID & "' WHERE BF_EMPNBR= " & EEID
    gdbAdoIhr001X.CommitTrans
End If

xFList = Get_Fields(gdbAdoIhr001X, "Term_HRBENS", "TERM_SEQ,BD_ID,BD_EMPNBR")
SQLQ = "INSERT INTO HRBENS (" & xFList & ", BD_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS BD_EMPNBR "
SQLQ = SQLQ & " FROM Term_HRBENS WHERE (Term_HRBENS.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans
If glbLinamar Then 'added by Bryan 17/Apr/2006 Ticket#10688
    SQLQ = "SELECT BD_BCODE FROM HRBENS WHERE BD_EMPNBR=" & CStr(EEID)
    rs.Open SQLQ, gdbAdoIhr001, adOpenDynamic, adLockOptimistic, adCmdText
    If rs.EOF = False Then
        Do
            rs("BD_BCODE") = CStr(Right(EEID, 3)) & Right(rs("BD_BCODE"), Len(rs("BD_BCODE")) - 3)
            rs.Update
            rs.MoveNext
        Loop Until rs.EOF
    End If
    rs.Close
End If

REIN_BENEFITS = True

Exit Function

REIN_BENEFITS_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_BENEFT", "REIN_BENEFT", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_COUNSEL(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_COUNSEL = False

On Error GoTo REIN_COUNSEL_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HR_COUNSEL", "TERM_SEQ,CL_ID,CL_EMPNBR")
SQLQ = "INSERT INTO HR_COUNSEL (" & xFList & ", CL_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS CL_EMPNBR "
SQLQ = SQLQ & " FROM Term_HR_COUNSEL WHERE (Term_HR_COUNSEL.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_COUNSEL = True

Exit Function

REIN_COUNSEL_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_COUNSEL", "Term_COUNSEL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_FOLLOW_UP(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_FOLLOW_UP = False

On Error GoTo REIN_FOLLOW_UP_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_FOLLOW_UP", "TERM_SEQ,EF_FOLLOWUP_ID,EF_EMPNBR")
SQLQ = "INSERT INTO HR_FOLLOW_UP (" & xFList & ", EF_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EF_EMPNBR "
SQLQ = SQLQ & " FROM Term_FOLLOW_UP WHERE (Term_FOLLOW_UP.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_FOLLOW_UP = True

Exit Function

REIN_FOLLOW_UP_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_FOLLOW_UP", "Term_FOLLOW_UP", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function
Function REIN_Profit_Sharing(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_Profit_Sharing = False

On Error GoTo REIN_PROFITSHARING_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_PROFIT_SHARING", "TERM_SEQ,PS_ID,PS_EMPNBR")
SQLQ = "INSERT INTO HR_PROFIT_SHARING (" & xFList & ", PS_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS PS_EMPNBR "
SQLQ = SQLQ & " FROM Term_PROFIT_SHARING WHERE (Term_PROFIT_SHARING.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_Profit_Sharing = True

Exit Function

REIN_PROFITSHARING_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_PROFITSHARING", "Term_PROFIT_SHARING", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Function
Function REIN_USERDEFINED(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_USERDEFINED = False

On Error GoTo REIN_USERDEFINED_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_USERDEFINE_TABLE", "TERM_SEQ,UD_ID,UD_EMPNBR")
SQLQ = "INSERT INTO HR_USERDEFINE_TABLE (" & xFList & ", UD_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS UD_EMPNBR "
SQLQ = SQLQ & " FROM Term_USERDEFINE_TABLE WHERE (Term_USERDEFINE_TABLE.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_USERDEFINED = True

Exit Function

REIN_USERDEFINED_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_USERDEFINED", "Term_USERDEFINE_TABLE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_EDUCSEM(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_EDUCSEM = False

On Error GoTo REIN_EDUCSEM_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HREDSEM", "TERM_SEQ,ES_ID,ES_EMPNBR")
SQLQ = "INSERT INTO HREDSEM (" & xFList & ", ES_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS ES_EMPNBR "
SQLQ = SQLQ & " FROM Term_HREDSEM WHERE (Term_HREDSEM.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_EDUCSEM = True

Exit Function

REIN_EDUCSEM_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_EDUCSEM", "term_EDUCSEM", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HREEO(EEID&, EESEQ&)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_HREEO = False

On Error GoTo REIN_HREEO_Err

'check if there is a record in HREEO
'Ticket #25969 Franks
SQLQ = "SELECT * FROM HREEO WHERE EO_EMPNBR = " & EEID& & " "
If rs.State <> 0 Then rs.Close
rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
If rs.EOF Then
    xFList = Get_Fields(gdbAdoIhr001X, "Term_HREEO", "TERM_SEQ,EO_EMPNBR")
    SQLQ = "INSERT INTO HREEO (" & xFList & ",EO_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EO_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HREEO WHERE (Term_HREEO.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001X.Execute SQLQ
End If

REIN_HREEO = True

Exit Function

REIN_HREEO_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HREEO", "Term_HREEO", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_EMP_FLAGS(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_EMP_FLAGS = False

On Error GoTo REIN_EMP_FLAGS_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HREMP_FLAGS", "TERM_SEQ,EF_ID,EF_EMPNBR")
SQLQ = "INSERT INTO HREMP_FLAGS (" & xFList & ", EF_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EF_EMPNBR "
SQLQ = SQLQ & " FROM Term_HREMP_FLAGS WHERE (Term_HREMP_FLAGS.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_EMP_FLAGS = True

Exit Function

REIN_EMP_FLAGS_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_EMP_FLAGS", "Term_HREMP_FLAGS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HREMPHIS(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_HREMPHIS = False

On Error GoTo REIN_HREMPHIS_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HREMPHIS", "TERM_SEQ,EE_ID,EE_EMPNBR")
SQLQ = "INSERT INTO HREMPHIS (" & xFList & ", EE_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EE_EMPNBR "
SQLQ = SQLQ & " FROM Term_HREMPHIS WHERE (Term_HREMPHIS.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_HREMPHIS = True

Exit Function

REIN_HREMPHIS_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HREMPHIS", "Term_HREMPHIS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HealthCost(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_HealthCost = False

On Error GoTo REIN_HealthCost_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HROHSCOS", "TERM_SEQ,CC_WCBC_ID,CC_EMPNBR")

SQLQ = "INSERT INTO HROHSCOS (" & xFList & ", CC_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS CC_EMPNBR "
SQLQ = SQLQ & " FROM Term_HROHSCOS WHERE (Term_HROHSCOS.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_HealthCost = True
Exit Function

REIN_HealthCost_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HealthCost", "REIN_HealthCost", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HealthSafety(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_HealthSafety = False

On Error GoTo REIN_HealthSafety_Err

If Not glbWFC Then 'wfc IncNum can be duplicated
    Call CheckDupIncNum(EEID&, EESEQ&)
End If

xFList = Get_Fields(gdbAdoIhr001X, "Term_HR_OCC_HEALTH_SAFETY", "TERM_SEQ,EC_ID,EC_EMPNBR")

SQLQ = "INSERT INTO HR_OCC_HEALTH_SAFETY  (" & xFList & ", EC_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EC_EMPNBR "
SQLQ = SQLQ & " FROM Term_HR_OCC_HEALTH_SAFETY WHERE (Term_HR_OCC_HEALTH_SAFETY.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_HealthSafety = True

Exit Function

REIN_HealthSafety_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HealthSafety", "REIN_HealthSafety", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function
Sub CheckDupIncNum(EEID&, EESEQ&)
Dim rsHS As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsOHSNBR As New ADODB.Recordset '
Dim SQLQ, xOHSnum

    SQLQ = "SELECT * FROM Term_HR_OCC_HEALTH_SAFETY WHERE TERM_SEQ = " & EESEQ&
    rsHS.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    SQLQ = "SELECT * FROM HROHSNBR "
    rsOHSNBR.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not rsHS.EOF
        SQLQ = "SELECT EC_EMPNBR,EC_CASE,EC_OCCDATE FROM HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & "WHERE EC_CASE = " & rsHS("EC_CASE")
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            xOHSnum = rsOHSNBR("OHSNBR")
            xOHSnum = xOHSnum + 1
            
            SQLQ = "UPDATE Term_OHS_CONTACT SET CT_Case = " & xOHSnum & " "
            SQLQ = SQLQ & "WHERE CT_Case = " & rsHS("EC_CASE") & " "
            SQLQ = SQLQ & "AND TERM_SEQ = " & EESEQ&
            gdbAdoIhr001X.Execute SQLQ
            
            SQLQ = "UPDATE Term_OHS_CORRECTIVE SET CR_Case = " & xOHSnum & " "
            SQLQ = SQLQ & "WHERE CR_Case = " & rsHS("EC_CASE") & " "
            SQLQ = SQLQ & "AND TERM_SEQ = " & EESEQ&
            gdbAdoIhr001X.Execute SQLQ
            
            SQLQ = "UPDATE Term_OHS_ROOT_CAUSES SET RC_Case = " & xOHSnum & " "
            SQLQ = SQLQ & "WHERE RC_Case = " & rsHS("EC_CASE") & " "
            SQLQ = SQLQ & "AND TERM_SEQ = " & EESEQ&
            gdbAdoIhr001X.Execute SQLQ
            
            SQLQ = "UPDATE Term_HROHSCOS SET CC_CASE = " & xOHSnum & " "
            SQLQ = SQLQ & "WHERE CC_CASE = " & rsHS("EC_CASE") & " "
            SQLQ = SQLQ & "AND TERM_SEQ = " & EESEQ&
            gdbAdoIhr001X.Execute SQLQ
            
            rsHS("EC_CASE") = xOHSnum
            rsHS.Update
            
            rsOHSNBR("OHSNBR") = xOHSnum
            rsOHSNBR.Update
        End If
        rsTemp.Close
        rsHS.MoveNext
    Loop
    rsOHSNBR.Close
    rsHS.Close
    If Not glbSQL And Not glbOracle Then Call Pause(1)
End Sub

Function REIN_OHS_CONTACT(EEID&, EESEQ&)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_OHS_CONTACT = False

On Error GoTo REIN_OHS_CONTACT_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_OHS_CONTACT", "TERM_SEQ,CT_ID,CT_Empnbr")

SQLQ = "INSERT INTO HR_OHS_CONTACT (" & xFList & ", CT_Empnbr) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS CT_Empnbr "
SQLQ = SQLQ & " FROM Term_OHS_CONTACT WHERE (Term_OHS_CONTACT.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_OHS_CONTACT = True

Exit Function

REIN_OHS_CONTACT_Err:
glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_OHS_CONTACT", "REIN_OHS_CONTACT", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_JOB(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_JOB = False

On Error GoTo REIN_JOB_Err


xFList = Get_Fields(gdbAdoIhr001X, "Term_JOB_HISTORY", "TERM_SEQ,JH_ID,JH_EMPNBR")
SQLQ = "INSERT INTO HR_JOB_HISTORY (" & xFList & ", JH_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS JH_EMPNBR "
SQLQ = SQLQ & " FROM Term_JOB_HISTORY WHERE (Term_JOB_HISTORY.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

'Hemu - Ticket #9802 - Jerry said to turn-off the Current Record flag when employee reinstated
'Franks 08/03/2012 Ticket #22392 This caused a problem at WFC, Jerry & Margaret asked to not use this for WFC
If Not glbWFC Then
    SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE (JH_CURRENT<>0 OR JH_TRK_CRS_RENEWAL<>0) AND JH_EMPNBR =" & EEID
    rsJH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    Do Until rsJH.EOF
        rsJH("JH_CURRENT") = 0
        rsJH("JH_TRK_CRS_RENEWAL") = 0
        If glbMulti Then
            SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE TERM_SEQ = " & EESEQ
            rsHrTerm.Open SQLQ, gdbAdoIhr001X, adOpenStatic, adLockPessimistic
            If Not rsHrTerm.EOF Then
                rsJH("JH_ENDDATE") = rsHrTerm("Term_DOT")
                rsJH("JH_ENDREAS") = rsHrTerm("Term_Reason")
            End If
            rsHrTerm.Close
        End If
        rsJH.Update
        rsJH.MoveNext
    Loop
    rsJH.Close
End If
'Hemu

'Friesens - Ticket #16189
'If glbCompSerial = "S/N - 2279W" Then
    xFList = Get_Fields(gdbAdoIhr001X, "Term_TEMP_WORK", "TERM_SEQ,TW_ID,TW_EMPNBR")
    SQLQ = "INSERT INTO HR_TEMP_WORK (" & xFList & ", TW_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS TW_EMPNBR "
    SQLQ = SQLQ & " FROM Term_TEMP_WORK WHERE (Term_TEMP_WORK.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001X.Execute SQLQ
'End If

'Turn-off the Current Record flag when employee reinstated
SQLQ = "SELECT * FROM HR_TEMP_WORK WHERE (TW_CURRENT<>0 OR TW_TRK_CRS_RENEWAL<>0) AND TW_EMPNBR =" & EEID
rsJH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
Do Until rsJH.EOF
    rsJH("TW_CURRENT") = 0
    rsJH("TW_TRK_CRS_RENEWAL") = 0
    If glbMulti Then
        SQLQ = "SELECT * FROM Term_HRTRMEMP WHERE TERM_SEQ = " & EESEQ
        rsHrTerm.Open SQLQ, gdbAdoIhr001X, adOpenStatic, adLockPessimistic
        If Not rsHrTerm.EOF Then
            rsJH("TW_ENDDATE") = rsHrTerm("Term_DOT")
            rsJH("TW_ENDREAS") = rsHrTerm("Term_Reason")
        End If
        rsHrTerm.Close
    End If
    rsJH.Update
    rsJH.MoveNext
Loop
rsJH.Close


REIN_JOB = True
Exit Function

REIN_JOB_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_JOB", "REIN_JOB", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_PERFORM(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
Dim rsPH As New ADODB.Recordset
REIN_PERFORM = False

On Error GoTo REIN_PERFORM_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_PERFORM_HISTORY", "TERM_SEQ,PH_ID,PH_EMPNBR")
SQLQ = "INSERT INTO HR_PERFORM_HISTORY (" & xFList & ", PH_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS PH_EMPNBR "
SQLQ = SQLQ & " FROM Term_PERFORM_HISTORY WHERE (Term_PERFORM_HISTORY.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

'Hemu - Ticket #9802 - Jerry said to turn-off the Current Record flag when employee reinstated
'Franks 08/03/2012 Ticket #22392 This caused a problem at WFC, Jerry & Mararet asked to not use this for WFC
If Not glbWFC Then
    SQLQ = "SELECT * FROM HR_PERFORM_HISTORY WHERE PH_CURRENT<>0 AND PH_EMPNBR =" & EEID
    rsPH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    Do Until rsPH.EOF
        rsPH("PH_CURRENT") = 0
        rsPH.Update
        rsPH.MoveNext
    Loop
    rsPH.Close
End If
'Hemu

If glbCompSerial = "S/N - 2279W" Then  'Friesens Corporation
    xFList = Get_Fields(gdbAdoIhr001X, "Term_PERFORM_FRIESEN", "TERM_SEQ,PH_ID,PH_EMPNBR")
    SQLQ = "INSERT INTO HR_PERFORM_FRIESEN (" & xFList & ", PH_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS PH_EMPNBR "
    SQLQ = SQLQ & " FROM Term_PERFORM_FRIESEN WHERE (Term_PERFORM_FRIESEN.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001X.Execute SQLQ
End If


REIN_PERFORM = True
Exit Function

REIN_PERFORM_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_PERFORM", "REIN_PERFORM", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Function

Function REIN_SALARY(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
Dim rsSH As New ADODB.Recordset
REIN_SALARY = False

On Error GoTo REIN_SALARY_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_SALARY_HISTORY", "TERM_SEQ,SH_ID,SH_EMPNBR")
SQLQ = "INSERT INTO HR_SALARY_HISTORY (" & xFList & ", SH_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS SH_EMPNBR "
SQLQ = SQLQ & " FROM Term_SALARY_HISTORY WHERE (Term_SALARY_HISTORY.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.BeginTrans
gdbAdoIhr001X.Execute SQLQ
gdbAdoIhr001X.CommitTrans

'Hemu - Ticket #9802 - Jerry said to turn-off the Current Record flag when employee reinstated
'Franks 08/03/2012 Ticket #22392 This caused a problem at WFC, Jerry & Margaret asked to not use this for WFC
If Not glbWFC Then
    SQLQ = "SELECT * FROM HR_SALARY_HISTORY WHERE SH_CURRENT<>0 AND SH_EMPNBR =" & EEID
    rsSH.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockPessimistic
    Do Until rsSH.EOF
        rsSH("SH_CURRENT") = 0
        rsSH.Update
        rsSH.MoveNext
    Loop
    rsSH.Close
End If
'Hemu

REIN_SALARY = True
Exit Function

REIN_SALARY_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_PERFORM", "REIN_PERFORM", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If


End Function
Function REIN_COMMENTS(EEID As Long, EESEQ As Long)   'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_COMMENTS = False

On Error GoTo REIN_COMMENTS_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_COMMENTS", "TERM_SEQ,CO_COMMENT_ID,CO_EMPNBR")

SQLQ = "INSERT INTO HR_COMMENTS (" & xFList & ", CO_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS CO_EMPNBR "
SQLQ = SQLQ & " FROM Term_COMMENTS WHERE (Term_COMMENTS.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_COMMENTS = True

Exit Function

REIN_COMMENTS_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_COMMENTS", "REIN_COMMENTS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function
Function REIN_OHS_CORRECTIVE(EEID As Long, EESEQ As Long)   'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_OHS_CORRECTIVE = False

On Error GoTo REIN_OHS_CORRECTIVE_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_OHS_CORRECTIVE", "TERM_SEQ,CR_ID,CR_EMPNBR")
SQLQ = "INSERT INTO HR_OHS_CORRECTIVE (" & xFList & ", CR_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS CR_EMPNBR "
SQLQ = SQLQ & " FROM Term_OHS_CORRECTIVE WHERE (Term_OHS_CORRECTIVE.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_OHS_CORRECTIVE = True
Exit Function

REIN_OHS_CORRECTIVE_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_OHS_CORRECTIVE", "REIN_OHS_CORRECTIVE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function
Function REIN_OHS_ROOT_CAUSES(EEID As Long, EESEQ As Long)    'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_OHS_ROOT_CAUSES = False

On Error GoTo REIN_OHS_ROOT_CAUSES_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_OHS_ROOT_CAUSES", "TERM_SEQ,RC_ID,RC_EMPNBR")
SQLQ = "INSERT INTO HR_OHS_ROOT_CAUSES (" & xFList & ", RC_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS RC_EMPNBR "
SQLQ = SQLQ & " FROM Term_OHS_ROOT_CAUSES WHERE (Term_OHS_ROOT_CAUSES.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ
REIN_OHS_ROOT_CAUSES = True
Exit Function

REIN_OHS_ROOT_CAUSES_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_OHS_ROOT_CAUSES", "REIN_OHS_ROOT_CAUSES", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_OHS_CLAIM_MEDICAL(EEID As Long, EESEQ As Long)    'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_OHS_CLAIM_MEDICAL = False

On Error GoTo REIN_OHS_CLAIM_MEDICAL_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_OHS_CLAIM_MEDICAL", "TERM_SEQ,EC_ID,EC_EMPNBR")
SQLQ = "INSERT INTO HR_OHS_CLAIM_MEDICAL (" & xFList & ", EC_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EC_EMPNBR "
SQLQ = SQLQ & " FROM Term_OHS_CLAIM_MEDICAL WHERE (Term_OHS_CLAIM_MEDICAL.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ
REIN_OHS_CLAIM_MEDICAL = True
Exit Function

REIN_OHS_CLAIM_MEDICAL_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_OHS_CLAIM_MEDICAL", "REIN_OHS_CLAIM_MEDICAL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_OHS_FORM7_SECTIONS(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_OHS_FORM7_SECTIONS = False

On Error GoTo REIN_OHS_FORM7_SECTIONS_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_OHS_FORM7_SECTIONS", "TERM_SEQ,F7_ID,F7_EMPNBR")
SQLQ = "INSERT INTO HR_OHS_FORM7_SECTIONS (" & xFList & ", F7_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS F7_EMPNBR "
SQLQ = SQLQ & " FROM Term_OHS_FORM7_SECTIONS WHERE (Term_OHS_FORM7_SECTIONS.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_OHS_FORM7_SECTIONS = True

Exit Function

REIN_OHS_FORM7_SECTIONS_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_OHS_FORM7_SECTIONS", "REIN_OHS_FORM7_SECTIONS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_OHS_FORM9(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_OHS_FORM9 = False

On Error GoTo REIN_OHS_FORM9_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_OHS_FORM9", "TERM_SEQ,F9_ID,F9_EMPNBR")
SQLQ = "INSERT INTO HR_OHS_FORM9 (" & xFList & ", F9_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS F9_EMPNBR "
SQLQ = SQLQ & " FROM Term_OHS_FORM9 WHERE (Term_OHS_FORM9.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_OHS_FORM9 = True

Exit Function

REIN_OHS_FORM9_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_OHS_FORM9", "REIN_OHS_FORM9", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_ENTHRS(EEID As Long, EESEQ As Long)    'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_ENTHRS = False

On Error GoTo REIN_ENTHRS_Err


xFList = Get_Fields(gdbAdoIhr001X, "Term_ENTHRS", "TERM_SEQ,HE_ID,HE_EMPNBR")
SQLQ = "INSERT INTO HRENTHRS (" & xFList & ", HE_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS HE_EMPNBR "
SQLQ = SQLQ & " FROM Term_ENTHRS WHERE (Term_ENTHRS.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_ENTHRS = True

Exit Function

REIN_ENTHRS_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_ENTHRS", "REIN_ENTHRS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_DOLENT(EEID As Long, EESEQ As Long)    'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_DOLENT = False

On Error GoTo REIN_DOLENT_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_DOLENT", "TERM_SEQ,DE_ENTITLE_ID,DE_EMPNBR")

SQLQ = "INSERT INTO HRDOLENT (" & xFList & ", DE_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DE_EMPNBR "
SQLQ = SQLQ & " FROM Term_DOLENT WHERE (Term_DOLENT.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_DOLENT = True

Exit Function

REIN_DOLENT_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_DOLENT", "REIN_DOLENT", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_DOLENT_ACTDTL(EEID As Long, EESEQ As Long)    'Ticket #28789 - Actual Amounts Details
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_DOLENT_ACTDTL = False

On Error GoTo REIN_DOLENT_ACTDTL_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_DOLENT_ACTDTL", "TERM_SEQ,DA_ENTITLE_ID,DA_EMPNBR")

SQLQ = "INSERT INTO HRDOLENT_ACTDTL (" & xFList & ", DA_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DA_EMPNBR "
SQLQ = SQLQ & " FROM Term_DOLENT_ACTDTL WHERE (Term_DOLENT_ACTDTL.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_DOLENT_ACTDTL = True

Exit Function

REIN_DOLENT_ACTDTL_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_DOLENT_ACTDTL", "REIN_DOLENT_ACTDTL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_EARN(EEID As Long, EESEQ As Long)    'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_EARN = False

On Error GoTo REIN_EARN_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_EARN", "TERM_SEQ,EARN_ID,EMPNBR")
SQLQ = "INSERT INTO HREARN (" & xFList & ", EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EMPNBR "
SQLQ = SQLQ & " FROM Term_EARN WHERE (Term_EARN.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_EARN = True
Exit Function

REIN_EARN_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_EARN", "REIN_EARN", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_EDU(EEID As Long, EESEQ As Long)   'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_EDU = False

On Error GoTo REIN_EDU_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_EDU", "TERM_SEQ,EU_ID,EU_EMPNBR")
SQLQ = "INSERT INTO HREDU (" & xFList & ", EU_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EU_EMPNBR "
SQLQ = SQLQ & " FROM Term_EDU WHERE (Term_EDU.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ
REIN_EDU = True
Exit Function

REIN_EDU_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_EDU", "REIN_EDU", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_DEPEND(EEID As Long, EESEQ As Long)   'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_DEPEND = False

On Error GoTo REIN_DEPEND_Err


xFList = Get_Fields(gdbAdoIhr001X, "Term_HRDEPEND", "TERM_SEQ,DP_ID,DP_EMPNBR")

SQLQ = "INSERT INTO HRDEPEND (" & xFList & ", DP_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DP_EMPNBR "
SQLQ = SQLQ & " FROM Term_HRDEPEND WHERE (Term_HRDEPEND.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ
REIN_DEPEND = True
Exit Function

REIN_DEPEND_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_DEPEND", "REIN_DEPEND", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HREMP_OTHER(EEID As Long, EESEQ As Long) 'Ticket #19488 Frank 11/29/10
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_HREMP_OTHER = False

On Error GoTo REIN_DEPEND_Err


xFList = Get_Fields(gdbAdoIhr001X, "Term_HREMP_OTHER", "TERM_SEQ,ER_ID,ER_EMPNBR")

SQLQ = "INSERT INTO HREMP_OTHER (" & xFList & ", ER_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS ER_EMPNBR "
SQLQ = SQLQ & " FROM Term_HREMP_OTHER WHERE (Term_HREMP_OTHER.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ
REIN_HREMP_OTHER = True
Exit Function

REIN_DEPEND_Err:
glbFrmCaption$ = "" ' "Terminate Emp "
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HREMP_OTHER", "REIN_HREMP_OTHER", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_COBRA(EEID As Long, EESEQ As Long)   'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_COBRA = False

On Error GoTo REIN_COBRA_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HRCOBRA", "TERM_SEQ,ID,EMPNBR")

SQLQ = "INSERT INTO HRCOBRA (" & xFList & ", EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EMPNBR "
SQLQ = SQLQ & " FROM Term_HRCOBRA WHERE (Term_HRCOBRA.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ
REIN_COBRA = True
Exit Function

REIN_COBRA_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_COBRA", "REIN_COBRA", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_EMPSKL(EEID As Long, EESEQ As Long)    'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_EMPSKL = False

On Error GoTo REIN_EMPSKL_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_EMPSKL", "TERM_SEQ,SE_ID,SE_EMPNBR")
SQLQ = "INSERT INTO HREMPSKL (" & xFList & ", SE_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS SE_EMPNBR "
SQLQ = SQLQ & " FROM Term_EMPSKL WHERE (Term_EMPSKL.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_EMPSKL = True
Exit Function

REIN_EMPSKL_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_EMPSKL", "REIN_EMPSKL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_SUCCESSION(EEID As Long, EESEQ As Long)    'George Apr4,2006 #10595
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_SUCCESSION = False

On Error GoTo REIN_SUCCESSION_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HR_SUCCESSION", "TERM_SEQ,EU_ID,EU_EMPNBR")
SQLQ = "INSERT INTO HR_SUCCESSION (" & xFList & ", EU_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EU_EMPNBR "
SQLQ = SQLQ & " FROM Term_HR_SUCCESSION WHERE (Term_HR_SUCCESSION.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_SUCCESSION = True
Exit Function

REIN_SUCCESSION_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_SUCCESSION", "REIN_SUCCESSION", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_LANGUAGE(EEID As Long, EESEQ As Long)    'George Apr4,2006 #10595
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_LANGUAGE = False

On Error GoTo REIN_LANGUAGE_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_HR_LANGUAGE", "TERM_SEQ,EL_ID,EL_EMPNBR")
SQLQ = "INSERT INTO HR_LANGUAGE (" & xFList & ", EL_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EL_EMPNBR "
SQLQ = SQLQ & " FROM Term_HR_LANGUAGE WHERE (Term_HR_LANGUAGE.TERM_SEQ= " & EESEQ & ")"
gdbAdoIhr001X.Execute SQLQ

REIN_LANGUAGE = True
Exit Function

REIN_LANGUAGE_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_LANGUAGE", "REIN_LANGUAGE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_TRADE(EEID As Long, EESEQ As Long)   'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_TRADE = False

On Error GoTo REIN_TRADE_Err


xFList = Get_Fields(gdbAdoIhr001X, "Term_TRADE", "TERM_SEQ,TD_ID,TD_EMPNBR")
SQLQ = "INSERT INTO HRTRADE (" & xFList & ", TD_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS TD_EMPNBR "
SQLQ = SQLQ & " FROM Term_TRADE WHERE (Term_TRADE.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_TRADE = True
Exit Function

REIN_TRADE_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_TRADE", "REIN_TRADE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_RSP(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_HRRSP = False

On Error GoTo REIN_HRRSP_Err


xFList = Get_Fields(gdbAdoIhr001X, "Term_HRRSP", "TERM_SEQ, RS_EMPNBR")
SQLQ = "INSERT INTO HRRSP (" & xFList & ", RS_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS RS_EMPNBR "
SQLQ = SQLQ & " FROM Term_HRRSP WHERE (Term_HRRSP.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_HRRSP = True
Exit Function

REIN_HRRSP_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRRSP", "REIN_HRRSP", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_LN_EMPSKL(EEID As Long, EESEQ As Long)   'JADDY 10/27/2005
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
REIN_LN_EMPSKL = False

On Error GoTo REIN_LN_EMPSKL_Err


xFList = Get_Fields(gdbAdoIhr001X, "LN_Term_EMPSKL", "TERM_SEQ,SE_ID,SE_EMPNBR")
SQLQ = "INSERT INTO LN_EMPSKL (" & xFList & ", SE_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS SE_EMPNBR "
SQLQ = SQLQ & " FROM LN_Term_EMPSKL WHERE (LN_Term_EMPSKL.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_LN_EMPSKL = True
Exit Function

REIN_LN_EMPSKL_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "LN_EMPSKL", "LN_EMPSKL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HR_PHOTO(EEID As Long, EESEQ As Long, Optional OldEmpNo)
    Dim SQLQ As String
    REIN_HR_PHOTO = False
    On Error GoTo REIN_HR_PHOTO_Err
    
    If Not IsMissing(OldEmpNo) Then
        SQLQ = "UPDATE HR_PHOTO SET PT_EMPNBR = " & EEID & " WHERE PT_EMPNBR = " & OldEmpNo
        gdbAdoIhr001X.Execute SQLQ
    End If
    
    REIN_HR_PHOTO = True
    
Exit Function

REIN_HR_PHOTO_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "HR_PHOTO", "HR_PHOTO", "Change")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_VACTIMEOFF_REQ(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_VACTIMEOFF_REQ = False

On Error GoTo REIN_VACTIMEOFF_REQ_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_VACTIMEOFF_REQ", "TERM_SEQ,VT_ID,VT_EMPNBR")

SQLQ = "INSERT INTO HR_VACTIMEOFF_REQ (" & xFList & ", VT_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS VT_EMPNBR "
SQLQ = SQLQ & " FROM Term_VACTIMEOFF_REQ WHERE (Term_VACTIMEOFF_REQ.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_VACTIMEOFF_REQ = True

Exit Function

REIN_VACTIMEOFF_REQ_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_VACTIMEOFF_REQ", "HR_VACTIMEOFF_REQ", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_VACTIMEOFF_REQ_ARCHIVE(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_VACTIMEOFF_REQ_ARCHIVE = False

On Error GoTo REIN_VACTIMEOFF_REQ_ARCHIVE_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_VACTIMEOFF_REQ_ARCHIVE", "TERM_SEQ,VT_ID,VT_EMPNBR")

SQLQ = "INSERT INTO HR_VACTIMEOFF_REQ_ARCHIVE (" & xFList & ", VT_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS VT_EMPNBR "
SQLQ = SQLQ & " FROM Term_VACTIMEOFF_REQ_ARCHIVE WHERE (Term_VACTIMEOFF_REQ_ARCHIVE.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_VACTIMEOFF_REQ_ARCHIVE = True

Exit Function

REIN_VACTIMEOFF_REQ_ARCHIVE_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_VACTIMEOFF_REQ_ARCHIVE", "HR_VACTIMEOFF_REQ_ARCHIVE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_VACTIMEOFF_REQ_WRK(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_VACTIMEOFF_REQ_WRK = False

On Error GoTo REIN_VACTIMEOFF_REQ_WRK_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_VACTIMEOFF_REQ_WRK", "TERM_SEQ,VT_WRK_ID,VT_EMPNBR")

SQLQ = "INSERT INTO HR_VACTIMEOFF_REQ_WRK (" & xFList & ", VT_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS VT_EMPNBR "
SQLQ = SQLQ & " FROM Term_VACTIMEOFF_REQ_WRK WHERE (Term_VACTIMEOFF_REQ_WRK.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_VACTIMEOFF_REQ_WRK = True

Exit Function

REIN_VACTIMEOFF_REQ_WRK_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_VACTIMEOFF_REQ_WRK", "HR_VACTIMEOFF_REQ_WRK", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_REQAUDIT(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_REQAUDIT = False

On Error GoTo REIN_REQAUDIT_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_REQAUDIT", "TERM_SEQ,RT_ID,RT_EMPNBR")

SQLQ = "INSERT INTO HR_REQAUDIT (" & xFList & ", RT_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS RT_EMPNBR "
SQLQ = SQLQ & " FROM Term_REQAUDIT WHERE (Term_REQAUDIT.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_REQAUDIT = True

Exit Function

REIN_REQAUDIT_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_REQAUDIT", "HR_REQAUDIT", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_TIMESHEET(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_TIMESHEET = False

On Error GoTo REIN_TIMESHEET_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_TIMESHEET", "TERM_SEQ,AD_ATT_ID,AD_EMPNBR")

SQLQ = "INSERT INTO HR_TIMESHEET (" & xFList & ", AD_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS AD_EMPNBR "
SQLQ = SQLQ & " FROM Term_TIMESHEET WHERE (Term_TIMESHEET.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_TIMESHEET = True

Exit Function

REIN_TIMESHEET_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_TIMESHEET", "HR_TIMESHEET", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_TIMESHEET_ARCHIVE(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

REIN_TIMESHEET_ARCHIVE = False

On Error GoTo REIN_TIMESHEET_ARCHIVE_Err

xFList = Get_Fields(gdbAdoIhr001X, "Term_TIMESHEET_ARCHIVE", "TERM_SEQ,AD_ATT_ID,AD_EMPNBR")

SQLQ = "INSERT INTO HR_TIMESHEET_ARCHIVE (" & xFList & ", AD_EMPNBR) "
SQLQ = SQLQ & in_SQL(glbIHRDB)
SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS AD_EMPNBR "
SQLQ = SQLQ & " FROM Term_TIMESHEET_ARCHIVE WHERE (Term_TIMESHEET_ARCHIVE.TERM_SEQ= " & EESEQ & ")"

gdbAdoIhr001X.Execute SQLQ

REIN_TIMESHEET_ARCHIVE = True

Exit Function

REIN_TIMESHEET_ARCHIVE_Err:
glbFrmCaption$ = "Terminate Emp"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_TIMESHEET_ARCHIVE", "HR_TIMESHEET_ARCHIVE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HRDOC_EMP(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_EMP = False

On Error GoTo REIN_HRDOC_EMP_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EMP", "TERM_SEQ,RE_ID,RE_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_EMP (" & xFList & ", RE_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS RE_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_EMP WHERE RE_TYPE='RESUME' AND (Term_HRDOC_EMP.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_EMP WHERE RE_TYPE='RESUME' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    If Not rsDocActive.EOF Then
        SQLQ = "SELECT * FROM HRDOC_EMP WHERE RE_TYPE='RESUME' AND RE_EMPNBR=" & EEID
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("RE_EMPNBR") = EEID
            rsDocTerm("RE_FILEEXT") = rsDocActive("RE_FILEEXT")
            rsDocTerm("RE_TYPE") = rsDocActive("RE_TYPE")
            rsDocTerm("RE_DOC") = rsDocActive("RE_DOC")
            rsDocTerm("RE_LDATE") = rsDocActive("RE_LDATE")
            rsDocTerm("RE_LTIME") = rsDocActive("RE_LTIME")
            rsDocTerm("RE_LUSER") = rsDocActive("RE_LUSER")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
    End If
    rsDocActive.Close
End If


SQLQ = "DELETE FROM Term_HRDOC_EMP "
SQLQ = SQLQ & "WHERE RE_TYPE='RESUME' AND TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_EMP = True

Exit Function

REIN_HRDOC_EMP_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_EMP", "REIN_HRDOC_EMP", "Insert - " & SQLQ)

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_HREMP_OTHER(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String

REIN_HRDOC_HREMP_OTHER = False

On Error GoTo REIN_HRDOC_HREMP_OTHER_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HREMP_OTHER", "TERM_SEQ,ER_ID,ER_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_HREMP_OTHER (" & xFList & ", ER_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS ER_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='OTHERINFO' AND (Term_HRDOC_HREMP_OTHER.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='OTHERINFO' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    If Not rsDocActive.EOF Then
        SQLQ = "SELECT * FROM HRDOC_EMP WHERE ER_TYPE='OTHERINFO' AND ER_EMPNBR=" & EEID
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("ER_EMPNBR") = EEID
            rsDocTerm("ER_FILEEXT") = rsDocActive("ER_FILEEXT")
            rsDocTerm("ER_TYPE") = rsDocActive("ER_TYPE")
            rsDocTerm("ER_DOC") = rsDocActive("ER_DOC")
            rsDocTerm("ER_LDATE") = rsDocActive("ER_LDATE")
            rsDocTerm("ER_LTIME") = rsDocActive("ER_LTIME")
            rsDocTerm("ER_LUSER") = rsDocActive("ER_LUSER")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
    End If
    rsDocActive.Close
End If


SQLQ = "DELETE FROM Term_HRDOC_HREMP_OTHER "
SQLQ = SQLQ & "WHERE ER_TYPE='OTHERINFO' AND TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_HREMP_OTHER = True

Exit Function

REIN_HRDOC_HREMP_OTHER_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_HREMP_OTHER", "REIN_HRDOC_HREMP_OTHER", "Insert - " & SQLQ)

Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_JOB_HISTORY(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_JOB_HISTORY = False

On Error GoTo REIN_HRDOC_JOB_HISTORY_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_JOB_HISTORY", "TERM_SEQ,DJ_ID,DJ_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_JOB_HISTORY (" & xFList & ", DJ_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DJ_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_JOB_HISTORY WHERE (Term_HRDOC_JOB_HISTORY.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_JOB_HISTORY WHERE DJ_TYPE='OFFER' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_JOB_HISTORY WHERE DJ_TYPE='OFFER' AND DJ_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND DJ_JOB = '" & rsDocActive("DJ_JOB") & "' "
        If IsDate(rsDocActive("DJ_SDATE")) Then
            SQLQ = SQLQ & "AND DJ_SDATE = " & Date_SQL(rsDocActive("DJ_SDATE")) & " "
        End If
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DJ_EMPNBR") = EEID
            rsDocTerm("DJ_SDATE") = rsDocActive("DJ_SDATE")
            rsDocTerm("DJ_JOB") = rsDocActive("DJ_JOB")
            rsDocTerm("DJ_DOC") = rsDocActive("DJ_DOC")
            rsDocTerm("DJ_FILEEXT") = rsDocActive("DJ_FILEEXT")
            rsDocTerm("DJ_TYPE") = rsDocActive("DJ_TYPE")
            rsDocTerm("DJ_LDATE") = rsDocActive("DJ_LDATE")
            rsDocTerm("DJ_LTIME") = rsDocActive("DJ_LTIME")
            rsDocTerm("DJ_LUSER") = rsDocActive("DJ_LUSER")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_JOB_HISTORY "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_JOB_HISTORY = True

Exit Function

REIN_HRDOC_JOB_HISTORY_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_JOB_HISTORY", "REIN_HRDOC_JOB_HISTORY", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_PERFORM_HISTORY(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_PERFORM_HISTORY = False

On Error GoTo REIN_HRDOC_PERFORM_HISTORY_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_PERFORM_HISTORY", "TERM_SEQ,DH_ID,DH_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_PERFORM_HISTORY (" & xFList & ", DH_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DH_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_PERFORM_HISTORY WHERE (Term_HRDOC_PERFORM_HISTORY.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_PERFORM_HISTORY WHERE DH_TYPE='PERFORMANCE' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_PERFORM_HISTORY WHERE DH_TYPE='PERFORMANCE' AND DH_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND DH_DOCKEY = " & rsDocActive("DH_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DH_EMPNBR") = DH_EMPNBR
            rsDocTerm("DH_JOB") = rsDocActive("DH_JOB")
            rsDocTerm("DH_PREVDATE") = rsDocActive("DH_PREVDATE")
            rsDocTerm("DH_DOC") = rsDocActive("DH_DOC")
            rsDocTerm("DH_FILEEXT") = rsDocActive("DH_FILEEXT")
            rsDocTerm("DH_TYPE") = rsDocActive("DH_TYPE")
            rsDocTerm("DH_LDATE") = rsDocActive("DH_LDATE")
            rsDocTerm("DH_LTIME") = rsDocActive("DH_LTIME")
            rsDocTerm("DH_LUSER") = rsDocActive("DH_LUSER")
            rsDocTerm("DH_DOCKEY") = rsDocActive("DH_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_PERFORM_HISTORY "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_PERFORM_HISTORY = True

Exit Function

REIN_HRDOC_PERFORM_HISTORY_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_PERFORM_HISTORY", "REIN_HRDOC_PERFORM_HISTORY", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_COMMENTS(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_COMMENTS = False

On Error GoTo REIN_HRDOC_COMMENTS_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_COMMENTS", "TERM_SEQ,DO_ID,DO_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_COMMENTS (" & xFList & ", DO_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DO_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_COMMENTS WHERE (Term_HRDOC_COMMENTS.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_COMMENTS WHERE DO_TYPE='COMMENTS' AND TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_COMMENTS WHERE DO_TYPE='COMMENTS' AND DO_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND DO_DOCKEY = " & rsDocActive("DO_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DO_EMPNBR") = EEID
            rsDocTerm("DO_COTYPE") = rsDocActive("DO_COTYPE")
            rsDocTerm("DO_EDATE") = rsDocActive("DO_EDATE")
            rsDocTerm("DO_DOC") = rsDocActive("DO_DOC")
            rsDocTerm("DO_FILEEXT") = rsDocActive("DO_FILEEXT")
            rsDocTerm("DO_TYPE") = rsDocActive("DO_TYPE")
            rsDocTerm("DO_LDATE") = rsDocActive("DO_LDATE")
            rsDocTerm("DO_LTIME") = rsDocActive("DO_LTIME")
            rsDocTerm("DO_LUSER") = rsDocActive("DO_LUSER")
            rsDocTerm("DO_DOCKEY") = rsDocActive("DO_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_COMMENTS "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_COMMENTS = True

Exit Function

REIN_HRDOC_COMMENTS_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_COMMENTS", "REIN_HRDOC_COMMENTS", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_HEALTH_SAFETY_2(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_HEALTH_SAFETY_2 = False

On Error GoTo REIN_HRDOC_HEALTH_SAFETY_2_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HEALTH_SAFETY_2", "TERM_SEQ,DE_ID,DE_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_HEALTH_SAFETY_2 (" & xFList & ", DE_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DE_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_HEALTH_SAFETY_2 WHERE (Term_HRDOC_HEALTH_SAFETY_2.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='INCIDENT' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='INCIDENT' AND DE_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND DE_CASE = " & rsDocActive("DE_CASE") & " "
        SQLQ = SQLQ & "AND DE_DOCNO = " & rsDocActive("DE_DOCNO") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DE_EMPNBR") = EEID
            rsDocTerm("DE_CASE") = rsDocActive("DE_CASE")
            rsDocTerm("DE_OCCDATE") = rsDocActive("DE_OCCDATE")
            rsDocTerm("DE_DOCNO") = rsDocActive("DE_DOCNO")
            rsDocTerm("DE_DOCDESC") = rsDocActive("DE_DOCDESC")
            rsDocTerm("DE_DOC") = rsDocActive("DE_DOC")
            rsDocTerm("DE_FILEEXT") = rsDocActive("DE_FILEEXT")
            rsDocTerm("DE_TYPE") = rsDocActive("DE_TYPE")
            rsDocTerm("DE_LDATE") = rsDocActive("DE_LDATE")
            rsDocTerm("DE_LTIME") = rsDocActive("DE_LTIME")
            rsDocTerm("DE_LUSER") = rsDocActive("DE_LUSER")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_HEALTH_SAFETY_2 "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_HEALTH_SAFETY_2 = True

Exit Function

REIN_HRDOC_HEALTH_SAFETY_2_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_HEALTH_SAFETY_2", "REIN_HRDOC_HEALTH_SAFETY_2", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_HEALTH_SAFETY(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_HEALTH_SAFETY = False

On Error GoTo REIN_HRDOC_HEALTH_SAFETY_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HEALTH_SAFETY", "TERM_SEQ,DE_ID,DE_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_HEALTH_SAFETY (" & xFList & ", DE_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DE_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_HEALTH_SAFETY WHERE (Term_HRDOC_HEALTH_SAFETY.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_HEALTH_SAFETY WHERE DE_TYPE='INCIDENT' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY WHERE DE_TYPE='INCIDENT' AND DE_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND DE_CASE = " & rsDocActive("DE_CASE") & " "
        SQLQ = SQLQ & "AND DE_DOCNO = " & rsDocActive("DE_DOCNO") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DE_EMPNBR") = EEID
            rsDocTerm("DE_CASE") = rsDocActive("DE_CASE")
            rsDocTerm("DE_OCCDATE") = rsDocActive("DE_OCCDATE")
            rsDocTerm("DE_DOCNO") = rsDocActive("DE_DOCNO")
            rsDocTerm("DE_DOCDESC") = rsDocActive("DE_DOCDESC")
            rsDocTerm("DE_DOC") = rsDocActive("DE_DOC")
            rsDocTerm("DE_FILEEXT") = rsDocActive("DE_FILEEXT")
            rsDocTerm("DE_TYPE") = rsDocActive("DE_TYPE")
            rsDocTerm("DE_LDATE") = rsDocActive("DE_LDATE")
            rsDocTerm("DE_LTIME") = rsDocActive("DE_LTIME")
            rsDocTerm("DE_LUSER") = rsDocActive("DE_LUSER")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_HEALTH_SAFETY "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_HEALTH_SAFETY = True

Exit Function

REIN_HRDOC_HEALTH_SAFETY_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_HEALTH_SAFETY", "REIN_HRDOC_HEALTH_SAFETY", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7 = False

On Error GoTo REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HEALTH_SAFETY_CONCERNSWF7", "TERM_SEQ,W7_ID,W7_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_HEALTH_SAFETY_CONCERNSWF7 (" & xFList & ", W7_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS W7_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE (Term_HRDOC_HEALTH_SAFETY_CONCERNSWF7.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='INJURYWF7' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='INJURYWF7' AND W7_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND W7_CASE = " & rsDocActive("W7_CASE") & " "
        SQLQ = SQLQ & "AND W7_DOCKEY = " & rsDocActive("W7_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("W7_EMPNBR") = EEID
            rsDocTerm("W7_CASE") = rsDocActive("W7_CASE")
            rsDocTerm("W7_OCCDATE") = rsDocActive("W7_OCCDATE")
            rsDocTerm("W7_DOCKEY") = rsDocActive("W7_DOCKEY")
            rsDocTerm("W7_DOCDESC") = rsDocActive("W7_DOCDESC")
            rsDocTerm("W7_DOC") = rsDocActive("W7_DOC")
            rsDocTerm("W7_FILEEXT") = rsDocActive("W7_FILEEXT")
            rsDocTerm("W7_TYPE") = rsDocActive("W7_TYPE")
            rsDocTerm("W7_LDATE") = rsDocActive("W7_LDATE")
            rsDocTerm("W7_LTIME") = rsDocActive("W7_LTIME")
            rsDocTerm("W7_LUSER") = rsDocActive("W7_LUSER")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_HEALTH_SAFETY_CONCERNSWF7 "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7 = True

Exit Function

REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7", "REIN_HRDOC_HEALTH_SAFETY_CONCERNSWF7", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_OHS_WRITTEN_OFFER(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_OHS_WRITTEN_OFFER = False

On Error GoTo REIN_HRDOC_OHS_WRITTEN_OFFER_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_OHS_WRITTEN_OFFER", "TERM_SEQ,F7_ID,F7_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_OHS_WRITTEN_OFFER (" & xFList & ", F7_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS F7_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_OHS_WRITTEN_OFFER WHERE (Term_HRDOC_OHS_WRITTEN_OFFER.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='INJURYWF7_WRITTENOFR' AND TERM_SEQ=" & EESEQ
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='INJURYWF7_WRITTENOFR' AND F7_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND F7_CASE = " & rsDocActive("F7_CASE") & " "
        SQLQ = SQLQ & "AND F7_DOCKEY = " & rsDocActive("F7_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("F7_EMPNBR") = EEID
            rsDocTerm("F7_CASE") = rsDocActive("F7_CASE")
            rsDocTerm("F7_OCCDATE") = rsDocActive("F7_OCCDATE")
            rsDocTerm("F7_DOCKEY") = rsDocActive("F7_DOCKEY")
            rsDocTerm("F7_DOCDESC") = rsDocActive("F7_DOCDESC")
            rsDocTerm("F7_DOC") = rsDocActive("F7_DOC")
            rsDocTerm("F7_FILEEXT") = rsDocActive("F7_FILEEXT")
            rsDocTerm("F7_TYPE") = rsDocActive("F7_TYPE")
            rsDocTerm("F7_LDATE") = rsDocActive("F7_LDATE")
            rsDocTerm("F7_LTIME") = rsDocActive("F7_LTIME")
            rsDocTerm("F7_LUSER") = rsDocActive("F7_LUSER")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_OHS_WRITTEN_OFFER "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_OHS_WRITTEN_OFFER = True

Exit Function

REIN_HRDOC_OHS_WRITTEN_OFFER_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_OHS_WRITTEN_OFFER", "REIN_HRDOC_OHS_WRITTEN_OFFER", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_EDSEM(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_EDSEM = False

On Error GoTo REIN_HRDOC_EDSEM_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EDSEM", "TERM_SEQ,ES_ID,ES_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_EDSEM (" & xFList & ", ES_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS ES_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_EDSEM WHERE (Term_HRDOC_EDSEM.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_EDSEM WHERE TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_EDSEM WHERE ES_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND ES_DOCKEY = " & rsDocActive("ES_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("ES_EMPNBR") = EEID
            rsDocTerm("ES_DOC") = rsDocActive("ES_DOC")
            rsDocTerm("ES_FILEEXT") = rsDocActive("ES_FILEEXT")
            rsDocTerm("ES_TYPE") = rsDocActive("ES_TYPE")
            rsDocTerm("ES_LDATE") = rsDocActive("ES_LDATE")
            rsDocTerm("ES_LTIME") = rsDocActive("ES_LTIME")
            rsDocTerm("ES_LUSER") = rsDocActive("ES_LUSER")
            rsDocTerm("ES_DOCKEY") = rsDocActive("ES_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_EDSEM "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_EDSEM = True

Exit Function

REIN_HRDOC_EDSEM_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_EDSEM", "REIN_HRDOC_EDSEM", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HRDOC_DOLENT(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_DOLENT = False

On Error GoTo REIN_HRDOC_DOLENT_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_DOLENT", "TERM_SEQ,DE_ID,DE_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_HRDOLENT (" & xFList & ", DE_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DE_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_DOLENT WHERE (Term_HRDOC_DOLENT.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_DOLENT WHERE TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_HRDOLENT WHERE DE_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND DE_DOCKEY = " & rsDocActive("DE_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DE_EMPNBR") = EEID
            rsDocTerm("DE_DOC") = rsDocActive("DE_DOC")
            rsDocTerm("DE_FILEEXT") = rsDocActive("DE_FILEEXT")
            rsDocTerm("DE_TYPE") = rsDocActive("DE_TYPE")
            rsDocTerm("DE_LDATE") = rsDocActive("DE_LDATE")
            rsDocTerm("DE_LTIME") = rsDocActive("DE_LTIME")
            rsDocTerm("DE_LUSER") = rsDocActive("DE_LUSER")
            rsDocTerm("DE_DOCKEY") = rsDocActive("DE_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_DOLENT "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_DOLENT = True

Exit Function

REIN_HRDOC_DOLENT_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_DOLENT", "REIN_HRDOC_DOLENT", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HRDOC_HREDU(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_HREDU = False

On Error GoTo REIN_HRDOC_HREDU_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HREDU", "TERM_SEQ,EU_ID,EU_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_HREDU (" & xFList & ", EU_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EU_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_HREDU WHERE (Term_HRDOC_HREDU.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_HREDU WHERE TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_HREDU WHERE EU_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND EU_DOCKEY = " & rsDocActive("EU_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("EU_EMPNBR") = EEID
            rsDocTerm("EU_DOC") = rsDocActive("EU_DOC")
            rsDocTerm("EU_FILEEXT") = rsDocActive("EU_FILEEXT")
            rsDocTerm("EU_TYPE") = rsDocActive("EU_TYPE")
            rsDocTerm("EU_LDATE") = rsDocActive("EU_LDATE")
            rsDocTerm("EU_LTIME") = rsDocActive("EU_LTIME")
            rsDocTerm("EU_LUSER") = rsDocActive("EU_LUSER")
            rsDocTerm("EU_DOCKEY") = rsDocActive("EU_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_HREDU "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_HREDU = True

Exit Function

REIN_HRDOC_HREDU_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_HREDU", "REIN_HRDOC_HREDU", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HRDOC_EDSEM_RETEST(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_EDSEM_RETEST = False

On Error GoTo REIN_HRDOC_EDSEM_RETEST_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EDSEM_RETEST", "TERM_SEQ,ES_ID,ES_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_EDSEM_RETEST (" & xFList & ", ES_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS ES_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_EDSEM_RETEST WHERE (Term_HRDOC_EDSEM_RETEST.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_EDSEM_RETEST WHERE TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_EDSEM_RETEST WHERE ES_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND ES_DOCKEY = " & rsDocActive("ES_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("ES_EMPNBR") = EEID
            rsDocTerm("ES_DOC") = rsDocActive("ES_DOC")
            rsDocTerm("ES_FILEEXT") = rsDocActive("ES_FILEEXT")
            rsDocTerm("ES_TYPE") = rsDocActive("ES_TYPE")
            rsDocTerm("ES_LDATE") = rsDocActive("ES_LDATE")
            rsDocTerm("ES_LTIME") = rsDocActive("ES_LTIME")
            rsDocTerm("ES_LUSER") = rsDocActive("ES_LUSER")
            rsDocTerm("ES_DOCKEY") = rsDocActive("ES_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_EDSEM_RETEST "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_EDSEM_RETEST = True

Exit Function

REIN_HRDOC_EDSEM_RETEST_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_EDSEM_RETEST", "REIN_HRDOC_EDSEM_RETEST", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HRDOC_COUNSEL(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String
REIN_HRDOC_COUNSEL = False

On Error GoTo REIN_HRDOC_COUNSEL_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_COUNSEL", "TERM_SEQ,DC_ID,DC_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_COUNSEL (" & xFList & ", DC_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS DC_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_COUNSEL WHERE (Term_HRDOC_COUNSEL.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_COUNSEL WHERE DC_TYPE='COUNSEL' AND TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_COUNSEL WHERE DC_TYPE='COUNSEL' AND DC_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND DC_DOCKEY = " & rsDocActive("DC_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("DC_EMPNBR") = EEID
            rsDocTerm("DC_CLTYPE") = rsDocActive("DC_CLTYPE")
            rsDocTerm("DC_COUDATE") = rsDocActive("DC_COUDATE")
            rsDocTerm("DC_DOC") = rsDocActive("DC_DOC")
            rsDocTerm("DC_FILEEXT") = rsDocActive("DC_FILEEXT")
            rsDocTerm("DC_TYPE") = rsDocActive("DC_TYPE")
            rsDocTerm("DC_LDATE") = rsDocActive("DC_LDATE")
            rsDocTerm("DC_LTIME") = rsDocActive("DC_LTIME")
            rsDocTerm("DC_LUSER") = rsDocActive("DC_LUSER")
            rsDocTerm("DC_DOCKEY") = rsDocActive("DC_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_COUNSEL "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_COUNSEL = True

Exit Function

REIN_HRDOC_COUNSEL_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_COUNSEL", "REIN_HRDOC_COUNSEL", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function REIN_HRDOC_TRADE(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String

REIN_HRDOC_TRADE = False

On Error GoTo REIN_HRDOC_TRADE_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_TRADE", "TERM_SEQ,TD_ID,TD_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_TRADE (" & xFList & ", TD_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS TD_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_TRADE WHERE (Term_HRDOC_TRADE.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_TRADE WHERE TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_TRADE WHERE TD_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND TD_DOCKEY = " & rsDocActive("TD_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("TD_EMPNBR") = EEID
            rsDocTerm("TD_CODE") = rsDocActive("TD_CODE")
            rsDocTerm("TD_BEGINDT") = rsDocActive("TD_BEGINDT")
            rsDocTerm("TD_DOC") = rsDocActive("TD_DOC")
            rsDocTerm("TD_FILEEXT") = rsDocActive("TD_FILEEXT")
            rsDocTerm("TD_TYPE") = rsDocActive("TD_TYPE")
            rsDocTerm("TD_LDATE") = rsDocActive("TD_LDATE")
            rsDocTerm("TD_LTIME") = rsDocActive("TD_LTIME")
            rsDocTerm("TD_LUSER") = rsDocActive("TD_LUSER")
            rsDocTerm("TD_DOCKEY") = rsDocActive("TD_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_TRADE "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_TRADE = True

Exit Function

REIN_HRDOC_TRADE_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_TRADE", "REIN_HRDOC_TRADE", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HRDOC_ATTENDANCE(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String

REIN_HRDOC_ATTENDANCE = False

On Error GoTo REIN_HRDOC_ATTENDANCE_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_ATTENDANCE", "TERM_SEQ,AD_ID,AD_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_ATTENDANCE (" & xFList & ", AD_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS AD_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_ATTENDANCE WHERE (Term_HRDOC_ATTENDANCE.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_ATTENDANCE WHERE TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_ATTENDANCE WHERE AD_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND AD_DOCKEY = " & rsDocActive("AD_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("AD_EMPNBR") = EEID
            rsDocTerm("AD_REASON") = rsDocActive("AD_REASON")
            rsDocTerm("AD_DOA") = rsDocActive("AD_DOA")
            rsDocTerm("AD_DOC") = rsDocActive("AD_DOC")
            rsDocTerm("AD_FILEEXT") = rsDocActive("AD_FILEEXT")
            rsDocTerm("AD_TYPE") = rsDocActive("AD_TYPE")
            rsDocTerm("AD_LDATE") = rsDocActive("AD_LDATE")
            rsDocTerm("AD_LTIME") = rsDocActive("AD_LTIME")
            rsDocTerm("AD_LUSER") = rsDocActive("AD_LUSER")
            rsDocTerm("AD_DOCKEY") = rsDocActive("AD_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_ATTENDANCE "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_ATTENDANCE = True

Exit Function

REIN_HRDOC_ATTENDANCE_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_ATTENDANCE", "REIN_HRDOC_ATTENDANCE", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function REIN_HRDOC_EMP_FLAGS(EEID As Long, EESEQ As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim rsJH As New ADODB.Recordset
Dim rsHrTerm As New ADODB.Recordset
Dim xFList As String

REIN_HRDOC_EMP_FLAGS = False

On Error GoTo REIN_HRDOC_EMP_FLAGS_Err

If glbSQL Then
    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EMP_FLAGS", "TERM_SEQ,EF_ID,EF_EMPNBR")
    SQLQ = "INSERT INTO HRDOC_EMP_FLAGS (" & xFList & ", EF_EMPNBR) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", " & EEID & " AS EF_EMPNBR "
    SQLQ = SQLQ & " FROM Term_HRDOC_EMP_FLAGS WHERE (Term_HRDOC_EMP_FLAGS.TERM_SEQ= " & EESEQ & ")"
    gdbAdoIhr001_DOC.BeginTrans
    gdbAdoIhr001_DOC.Execute SQLQ
    gdbAdoIhr001_DOC.CommitTrans
Else 'Oracle
    Dim rsDocActive As New ADODB.Recordset
    Dim rsDocTerm As New ADODB.Recordset
    SQLQ = "SELECT * FROM TERM_HRDOC_EMP_FLAGS WHERE TERM_SEQ=" & EESEQ & " "
    rsDocActive.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
    Do While Not rsDocActive.EOF
        SQLQ = "SELECT * FROM HRDOC_EMP_FLAGS WHERE EF_EMPNBR=" & EEID & " "
        SQLQ = SQLQ & "AND EF_DOCKEY = " & rsDocActive("EF_DOCKEY") & " "
        rsDocTerm.Open SQLQ, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        If rsDocTerm.EOF Then
            rsDocTerm.AddNew
            rsDocTerm("EF_EMPNBR") = EEID
            rsDocTerm("EF_FLAG") = rsDocActive("EF_FLAG")
            rsDocTerm("EF_FLAGDTE") = rsDocActive("EF_FLAGDTE")
            rsDocTerm("EF_DOC") = rsDocActive("EF_DOC")
            rsDocTerm("EF_FILEEXT") = rsDocActive("EF_FILEEXT")
            rsDocTerm("EF_TYPE") = rsDocActive("EF_TYPE")
            rsDocTerm("EF_LDATE") = rsDocActive("EF_LDATE")
            rsDocTerm("EF_LTIME") = rsDocActive("EF_LTIME")
            rsDocTerm("EF_LUSER") = rsDocActive("EF_LUSER")
            rsDocTerm("EF_DOCKEY") = rsDocActive("EF_DOCKEY")
            rsDocTerm.Update
        End If
        rsDocTerm.Close
        rsDocActive.MoveNext
    Loop
    rsDocActive.Close
End If

SQLQ = "DELETE FROM Term_HRDOC_EMP_FLAGS "
SQLQ = SQLQ & "WHERE TERM_SEQ=" & EESEQ

gdbAdoIhr001_DOC.Execute SQLQ

REIN_HRDOC_EMP_FLAGS = True

Exit Function

REIN_HRDOC_EMP_FLAGS_Err:

glbFrmCaption$ = "Rehire Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "REIN_HRDOC_EMP_FLAGS", "REIN_HRDOC_EMP_FLAGS", "Insert - " & SQLQ)
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_COMMENTS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_COMMENTS = False

On Error GoTo TERM_COMMENTS_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_COMMENTS", "CO_COMMENT_ID")
SQLQ = "INSERT INTO Term_COMMENTS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_COMMENTS "
SQLQ = SQLQ & "WHERE (HR_COMMENTS.CO_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_COMMENTS = True

Exit Function

TERM_COMMENTS_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_Comments", "Term_Comments", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_DOLENT(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_DOLENT = False

On Error GoTo TERM_DOLENT_Err
xFList = Get_Fields(gdbAdoIhr001, "HRDOLENT", "DE_ENTITLE_ID")
SQLQ = "INSERT INTO Term_DOLENT (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRDOLENT "
SQLQ = SQLQ & "WHERE (HRDOLENT.DE_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_DOLENT = True

Exit Function

TERM_DOLENT_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_DOLENT", "Term_DOLENT", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_DOLENT_ACTDTL(EEID As Long)   'Ticket #28789 - Actual Amounts Details
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_DOLENT_ACTDTL = False

On Error GoTo TERM_DOLENT_ACTDTL_Err

xFList = Get_Fields(gdbAdoIhr001, "HRDOLENT_ACTDTL", "DA_ENTITLE_ID")

SQLQ = "INSERT INTO Term_DOLENT_ACTDTL (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRDOLENT_ACTDTL "
SQLQ = SQLQ & "WHERE (HRDOLENT_ACTDTL.DA_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_DOLENT_ACTDTL = True

Exit Function

TERM_DOLENT_ACTDTL_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_DOLENT_ACTDTL", "Term_DOLENT_ACTDTL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_ENTHRS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_ENTHRS = False

On Error GoTo TERM_ENTHRS_Err
xFList = Get_Fields(gdbAdoIhr001, "HRENTHRS", "HE_ID")
SQLQ = "INSERT INTO Term_ENTHRS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRENTHRS "
SQLQ = SQLQ & "WHERE (HRENTHRS.HE_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_ENTHRS = True

Exit Function

TERM_ENTHRS_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_ENTHRS", "Term_ENTHRS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_EARN(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EARN = False

On Error GoTo TERM_EARN_Err

xFList = Get_Fields(gdbAdoIhr001, "HREARN", "EARN_ID")
SQLQ = "INSERT INTO Term_EARN (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREARN "
SQLQ = SQLQ & "WHERE (HREARN.EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_EARN = True

Exit Function

TERM_EARN_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_EARN", "TERM_EARN", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_EDU(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EDU = False

On Error GoTo TERM_EDU_Err

xFList = Get_Fields(gdbAdoIhr001, "HREDU", "EU_ID")
SQLQ = "INSERT INTO Term_EDU (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREDU "
SQLQ = SQLQ & "WHERE (HREDU.EU_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ
TERM_EDU = True

Exit Function

TERM_EDU_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_EDU", "TERM_EDU", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_DEPEND(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_DEPEND = False

On Error GoTo TERM_DEPEND_Err
If glbWFC And IsDate(glbChgBenTermDate) Then
    xFList = Get_Fields(gdbAdoIhr001, "HRDEPEND", "DP_ID,DP_EDATE")
    SQLQ = "INSERT INTO Term_HRDEPEND (" & xFList & ", TERM_SEQ,DP_EDATE) "
Else
    xFList = Get_Fields(gdbAdoIhr001, "HRDEPEND", "DP_ID")
    SQLQ = "INSERT INTO Term_HRDEPEND (" & xFList & ", TERM_SEQ) "
End If
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
If glbWFC And IsDate(glbChgBenTermDate) Then
    SQLQ = SQLQ & "," & Date_SQL(glbChgBenTermDate) & " As DP_EDATE "
End If
SQLQ = SQLQ & "FROM HRDEPEND "
SQLQ = SQLQ & "WHERE (HRDEPEND.DP_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ
TERM_DEPEND = True

Exit Function

TERM_DEPEND_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRDEPEND", "TERM_HRDEPEND", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_COBRA(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_COBRA = False

On Error GoTo TERM_COBRA_Err

xFList = Get_Fields(gdbAdoIhr001, "HRCOBRA", "ID")
SQLQ = "INSERT INTO Term_HRCOBRA (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRCOBRA "
SQLQ = SQLQ & "WHERE (HRCOBRA.EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ
TERM_COBRA = True

Exit Function

TERM_COBRA_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRCOBRA", "TERM_HRCOBRA", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_EMPSKL(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EMPSKL = False

On Error GoTo TERM_EMPSKL_Err
xFList = Get_Fields(gdbAdoIhr001, "HREMPSKL", "SE_ID")
SQLQ = "INSERT INTO Term_EMPSKL (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREMPSKL "
SQLQ = SQLQ & "WHERE (HREMPSKL.SE_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_EMPSKL = True

Exit Function

TERM_EMPSKL_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_EMPSKL", "TERM_EMPSKL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_SUCCESSION(EEID As Long) 'George Apr 4,2006 #10595
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_SUCCESSION = False

On Error GoTo TERM_SUCCESSION_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_SUCCESSION", "EU_ID")
SQLQ = "INSERT INTO Term_HR_SUCCESSION (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_SUCCESSION "
SQLQ = SQLQ & "WHERE (HR_SUCCESSION.EU_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_SUCCESSION = True

Exit Function

TERM_SUCCESSION_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HR_SUCCESSION", "TERM_HR_SUCCESSION", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_LANGUAGE(EEID As Long) 'George Apr 4,2006 #10595
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_LANGUAGE = False

On Error GoTo TERM_LANGUAGE_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_LANGUAGE", "EL_ID")
SQLQ = "INSERT INTO Term_HR_LANGUAGE (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_LANGUAGE "
SQLQ = SQLQ & "WHERE (HR_LANGUAGE.EL_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_LANGUAGE = True

Exit Function

TERM_LANGUAGE_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HR_LANGUAGE", "TERM_HR_LANGUAGE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_TRADE(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_TRADE = False

On Error GoTo TERM_TRADE_Err
xFList = Get_Fields(gdbAdoIhr001, "HRTRADE", "TD_ID")
SQLQ = "INSERT INTO Term_TRADE (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRTRADE "
SQLQ = SQLQ & "WHERE (HRTRADE.TD_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_TRADE = True

Exit Function

TERM_TRADE_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_TRADE", "TERM_TRADE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_OHS_Corrective(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_OHS_Corrective = False

On Error GoTo TERM_OHS_Corrective_Err
xFList = Get_Fields(gdbAdoIhr001, "HR_OHS_CORRECTIVE", "CR_ID")
SQLQ = "INSERT INTO Term_OHS_CORRECTIVE (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OHS_CORRECTIVE "
SQLQ = SQLQ & "WHERE (HR_OHS_CORRECTIVE.CR_Empnbr=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_OHS_Corrective = True

Exit Function

TERM_OHS_Corrective_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_OHS_Corrective", "Term_OHS_Corrective", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function Term_OHS_ROOT_CAUSES(EEID As Long) 'FRANK 4/5/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
Term_OHS_ROOT_CAUSES = False

On Error GoTo TERM_OHS_ROOT_CAUSES_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_OHS_ROOT_CAUSES", "RC_ID")
SQLQ = "INSERT INTO Term_OHS_ROOT_CAUSES (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OHS_ROOT_CAUSES "
SQLQ = SQLQ & "WHERE (HR_OHS_ROOT_CAUSES.RC_Empnbr=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

Term_OHS_ROOT_CAUSES = True

Exit Function

TERM_OHS_ROOT_CAUSES_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_OHS_ROOT_CAUSES", "Term_OHS_ROOT_CAUSES", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function Term_OHS_CLAIM_MEDICAL(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
Term_OHS_CLAIM_MEDICAL = False

On Error GoTo Term_OHS_CLAIM_MEDICAL_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_OHS_CLAIM_MEDICAL", "EC_ID")
SQLQ = "INSERT INTO Term_OHS_CLAIM_MEDICAL (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OHS_CLAIM_MEDICAL "
SQLQ = SQLQ & "WHERE (HR_OHS_CLAIM_MEDICAL.EC_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

Term_OHS_CLAIM_MEDICAL = True

Exit Function

Term_OHS_CLAIM_MEDICAL_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_OHS_CLAIM_MEDICAL", "Term_OHS_CLAIM_MEDICAL", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function Term_OHS_FORM7_SECTIONS(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

Term_OHS_FORM7_SECTIONS = False

On Error GoTo Term_OHS_FORM7_SECTIONS_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_OHS_FORM7_SECTIONS", "F7_ID")
SQLQ = "INSERT INTO Term_OHS_FORM7_SECTIONS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OHS_FORM7_SECTIONS "
SQLQ = SQLQ & "WHERE (HR_OHS_FORM7_SECTIONS.F7_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

Term_OHS_FORM7_SECTIONS = True

Exit Function

Term_OHS_FORM7_SECTIONS_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_OHS_FORM7_SECTIONS", "Term_OHS_FORM7_SECTIONS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function Term_OHS_FORM9(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

Term_OHS_FORM9 = False

On Error GoTo Term_OHS_FORM9_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_OHS_FORM9", "F9_ID")
SQLQ = "INSERT INTO Term_OHS_FORM9 (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_OHS_FORM9 "
SQLQ = SQLQ & "WHERE (HR_OHS_FORM9.F9_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

Term_OHS_FORM9 = True

Exit Function

Term_OHS_FORM9_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_OHS_FORM9", "Term_OHS_FORM9", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_RSP(EEID As Long) 'FRANK 12/21/2000
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_RSP = False

On Error GoTo TERM_RSP_Err
xFList = Get_Fields(gdbAdoIhr001, "HRRSP", "")
SQLQ = "INSERT INTO TERM_HRRSP (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRRSP "
SQLQ = SQLQ & "WHERE (HRRSP.RS_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_RSP = True

Exit Function

TERM_RSP_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRRSP", "TERM_HRRSP", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_LN_EMPSKL(EEID As Long) 'Jaddy 10/27/2005
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_LN_EMPSKL = False

On Error GoTo TERM_RSP_Err
xFList = Get_Fields(gdbAdoIhr001, "LN_EMPSKL", "SE_ID")

SQLQ = "INSERT INTO LN_TERM_EMPSKL (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM LN_EMPSKL "
SQLQ = SQLQ & "WHERE (LN_EMPSKL.SE_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_LN_EMPSKL = True

Exit Function

TERM_RSP_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "TERM_HRRSP", "TERM_HRRSP", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_EMP_FLAGS(EEID As Long) 'George Apr 4,2006 #10595
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_EMP_FLAGS = False

On Error GoTo TERM_SUCCESSION_Err
xFList = Get_Fields(gdbAdoIhr001, "HREMP_FLAGS", "EF_ID")
SQLQ = "INSERT INTO Term_HREMP_FLAGS (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HREMP_FLAGS "
SQLQ = SQLQ & "WHERE (HREMP_FLAGS.EF_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_EMP_FLAGS = True

Exit Function

TERM_SUCCESSION_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HREMP_FLAGS", "Term_HREMP_FLAGS", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_GLDIST(EEID As Long) 'George Apr 4,2006 #10595
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String
TERM_GLDIST = False

On Error GoTo TERM_SUCCESSION_Err
xFList = Get_Fields(gdbAdoIhr001, "HRGLDIST", "GL_ID")
SQLQ = "INSERT INTO Term_HRGLDIST (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT  " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HRGLDIST "
SQLQ = SQLQ & "WHERE (HRGLDIST.GL_EMPNBR=" & EEID & " )"

gdbAdoIhr001.Execute SQLQ

TERM_GLDIST = True

Exit Function

TERM_SUCCESSION_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_HRGLDIST", "Term_HRGLDIST", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If
End Function

Function TERM_VACTIMEOFF_REQ(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_VACTIMEOFF_REQ = False

On Error GoTo TERM_VACTIMEOFF_REQ_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_VACTIMEOFF_REQ", "VT_ID")
SQLQ = "INSERT INTO Term_VACTIMEOFF_REQ (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_VACTIMEOFF_REQ "
SQLQ = SQLQ & "WHERE (HR_VACTIMEOFF_REQ.VT_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_VACTIMEOFF_REQ = True

Exit Function

TERM_VACTIMEOFF_REQ_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_VACTIMEOFF_REQ", "Term_VACTIMEOFF_REQ", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_VACTIMEOFF_REQ_ARCH(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_VACTIMEOFF_REQ_ARCH = False

On Error GoTo TERM_VACTIMEOFF_REQ_ARCH_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_VACTIMEOFF_REQ_ARCHIVE", "VT_ID")
SQLQ = "INSERT INTO Term_VACTIMEOFF_REQ_ARCHIVE (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_VACTIMEOFF_REQ_ARCHIVE "
SQLQ = SQLQ & "WHERE (HR_VACTIMEOFF_REQ_ARCHIVE.VT_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_VACTIMEOFF_REQ_ARCH = True

Exit Function

TERM_VACTIMEOFF_REQ_ARCH_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_VACTIMEOFF_REQ_ARCHIVE", "Term_VACTIMEOFF_REQ_ARCHIVE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_VACTIMEOFF_REQ_WRK(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_VACTIMEOFF_REQ_WRK = False

On Error GoTo TERM_VACTIMEOFF_REQ_WRK_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_VACTIMEOFF_REQ_WRK", "VT_WRK_ID")
SQLQ = "INSERT INTO Term_VACTIMEOFF_REQ_WRK (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_VACTIMEOFF_REQ_WRK "
SQLQ = SQLQ & "WHERE (HR_VACTIMEOFF_REQ_WRK.VT_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_VACTIMEOFF_REQ_WRK = True

Exit Function

TERM_VACTIMEOFF_REQ_WRK_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_VACTIMEOFF_REQ_WRK", "Term_VACTIMEOFF_REQ_WRK", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_REAUDIT(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_REAUDIT = False

On Error GoTo TERM_REAUDIT_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_REQAUDIT", "RT_ID")
SQLQ = "INSERT INTO Term_REQAUDIT (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_REQAUDIT "
SQLQ = SQLQ & "WHERE (HR_REQAUDIT.RT_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_REAUDIT = True

Exit Function

TERM_REAUDIT_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_REQAUDIT", "Term_REQAUDIT", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_TIMESHEET(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_TIMESHEET = False

On Error GoTo TERM_TIMESHEET_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_TIMESHEET", "AD_ATT_ID")
SQLQ = "INSERT INTO Term_TIMESHEET (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_TIMESHEET "
SQLQ = SQLQ & "WHERE (HR_TIMESHEET.AD_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_TIMESHEET = True

Exit Function

TERM_TIMESHEET_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_TIMESHEET", "Term_TIMESHEET", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Function TERM_TIMESHEET_ARCH(EEID As Long)
Dim SQLQ As String
Dim iRow As Integer, Msg As String
Dim xFList As String

TERM_TIMESHEET_ARCH = False

On Error GoTo TERM_TIMESHEET_ARCH_Err

xFList = Get_Fields(gdbAdoIhr001, "HR_TIMESHEET_ARCHIVE", "AD_ATT_ID")
SQLQ = "INSERT INTO Term_TIMESHEET_ARCHIVE (" & xFList & ", TERM_SEQ) "
SQLQ = SQLQ & in_SQL(glbIHRAUDIT)
SQLQ = SQLQ & "SELECT " & xFList & ","
SQLQ = SQLQ & glbTERM_Seq & " As TERM_SEQ "
SQLQ = SQLQ & "FROM HR_TIMESHEET_ARCHIVE "
SQLQ = SQLQ & "WHERE (HR_TIMESHEET_ARCHIVE.AD_EMPNBR=" & EEID & " )"
gdbAdoIhr001.Execute SQLQ

TERM_TIMESHEET_ARCH = True

Exit Function

TERM_TIMESHEET_ARCH_Err:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "Terminate Employee"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Term_TIMESHEET_ARCHIVE", "Term_TIMESHEET_ARCHIVE", "Insert")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

End Function

Public Function GetTransDivTReason(xCode)
Dim rsTran As New ADODB.Recordset
Dim SQLQ As String
Dim xFirstVal As String
Dim xFinal As String

    xFirstVal = xCode
    xFinal = "'*'"
    If Not Len(xFirstVal) = 0 Then
        SQLQ = "SELECT * FROM "
        SQLQ = SQLQ & "HR_TERMCAUSE_LINK "
        SQLQ = SQLQ & "WHERE RL_FIRSTCODE ='" & xFirstVal & "' "
        rsTran.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsTran.EOF
            xFinal = xFinal & ",'" & rsTran("RL_SECONDCODE") & "'"
            rsTran.MoveNext
        Loop
    End If
    If xFinal = "'*'" Then
        xFinal = ""
    End If
    GetTransDivTReason = xFinal
End Function
