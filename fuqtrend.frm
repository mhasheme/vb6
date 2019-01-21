VERSION 5.00
Begin VB.Form frmUQuarterEnd 
   Caption         =   "Quarter End"
   ClientHeight    =   7092
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7092
   ScaleWidth      =   8760
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmdQTR 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   3
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label lblNewRec 
      Caption         =   "New Record Date"
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   2040
      Width           =   3375
   End
   Begin VB.Label lbl1 
      Caption         =   "BTI Only"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblTo 
      Caption         =   "To"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblDRdesc 
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   7
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label lblDRdesc 
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblDateRange 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date Range:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblSelCri 
      Caption         =   "Selection Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblQTRNo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Quarter"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   630
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Label lblYear 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
      Height          =   255
      Left            =   630
      TabIndex        =   0
      Top             =   630
      Width           =   975
   End
End
Attribute VB_Name = "frmUQuarterEnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub SetDateRange()
    If Len(txtYear) = 4 Then
        If IsNumeric(txtYear) Then
            If glbFormCaption = "Quarter End" Then
                If cmdQTR.Text = "1" Then
                    lblDRdesc(0) = GetMonth("Jan") & " 1, " & txtYear
                    lblDRdesc(1) = GetMonth("Mar") & " 31, " & txtYear
                End If
                If cmdQTR.Text = "2" Then
                    lblDRdesc(0) = GetMonth("Apr") & " 1, " & txtYear
                    lblDRdesc(1) = GetMonth("Jun") & " 30, " & txtYear
                End If
                If cmdQTR.Text = "3" Then
                    lblDRdesc(0) = GetMonth("Jul") & " 1, " & txtYear
                    lblDRdesc(1) = GetMonth("Sep") & " 30, " & txtYear
                End If
                If cmdQTR.Text = "4" Then
                    lblDRdesc(0) = GetMonth("Oct") & " 1, " & txtYear
                    lblDRdesc(1) = GetMonth("Dec") & " 31, " & txtYear
                End If
            Else
                lblDRdesc(0) = GetMonth("Jan") & " 1, " & txtYear
                lblDRdesc(1) = GetMonth("Dec") & " 31, " & txtYear
                
                If glbFormCaption = "Year End Carryover" Then
                    lblNewRec.Caption = "The Date of Year End Carryover is " & "Jan 1, " & txtYear + 1
                End If
                If glbFormCaption = "Year End Reduction For Non BD" Then
                    lblNewRec.Caption = "The Date of Year End Reduction For Non BD is " & "Jan 1, " & txtYear + 1
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdQTR_Click()
Call SetDateRange
End Sub

Public Sub cmdModify_Click()
Dim DgDef, Title$, Msg$, Response%, SQLQ
Dim rsTemp As New ADODB.Recordset
Dim xFlag As Boolean, xTDate
If Not chekQtrEnd() Then
  Exit Sub
End If

'If Not gSec_Upd_Earnings Then
'    MsgBox "You Do Not Have Authority For This Transaction"
'    Exit Sub
'End If
If glbFormCaption = "Quarter End" Then
    'Check if the Quarterly Reduction has been done for this Quarter
    xTDate = CVDate(lblDRdesc(1))
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (1=1) "
    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
    SQLQ = SQLQ & "AND (AD_REASON='QRED') "
    xFlag = False
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xFlag = True
    End If
    rsTemp.Close

    Title$ = "Quarterly Reduction"
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Msg$ = ""
    If xFlag Then
        Msg$ = Msg$ & "The Quarter " & cmdQTR.Text & " of " & txtYear & " Reduction has been finished. " & Chr(10)
    End If
    Msg$ = Msg$ & "Are You Sure You Want To Do The Quarterly Reduction?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    Call modQtrEnd
End If

If glbFormCaption = "Year End Carryover" Then
    'Check if the Year End Carryover has been done for this Year
    xTDate = CVDate(GetMonth("Jan") & " 1," & txtYear + 1)
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (1=1) "
    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
    SQLQ = SQLQ & "AND (AD_REASON='2100' OR AD_REASON='2200') "
    xFlag = False
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xFlag = True
    End If
    rsTemp.Close
    
    Title$ = "Year End Carryover"
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Msg$ = ""
    If xFlag Then
        Msg$ = Msg$ & "The " & txtYear & " Year Carryover has been finished. " & Chr(10)
    End If
    Msg$ = Msg$ & "Are You Sure You Want To Do The Year End Carryover?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    Call modYTDcarryover
End If

If glbFormCaption = "Year End Reduction For Non BD" Then
    'Check if the Year End Reduction For Non BD has been done for this Year
    xTDate = CVDate(GetMonth("Jan") & " 1," & txtYear + 1)
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (1=1) "
    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
    SQLQ = SQLQ & "AND (AD_REASON='2300' OR AD_REASON='2400') "
    SQLQ = SQLQ & "AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION='HRLY' AND NOT (ED_DIV='BD')) "
    xFlag = False
    If rsTemp.State <> 0 Then rsTemp.Close
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        xFlag = True
    End If
    rsTemp.Close
    
    Title$ = "Year End Reduction For Non BD"
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Msg$ = ""
    If xFlag Then
        Msg$ = Msg$ & "The " & txtYear & " Year End Reduction For Non BD has been finished. " & Chr(10)
    End If
    Msg$ = Msg$ & "Are You Sure You Want To Do The Year End Reduction For Non BD?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    Call modYTD_Reduction_NonBD
End If
End Sub
Private Function modYTD_Reduction_NonBD()
Dim rsTAtt As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsMain As New ADODB.Recordset
Dim SQLQ, xNum, xCode, xSec, xEmpNo
Dim xUnexFlag, xEmlFlag, xUnexVal, xExcuVal, xEmlVal, glbDiv, glbYear, xYear, glbPointType
Dim I, xNTot, xDOA
Dim xFDate, xTDate
    xFDate = CVDate(GetMonth("Jan") & " 1," & txtYear + 1)
    xTDate = CVDate(GetMonth("Dec") & " 31," & txtYear + 1)
    xYear = Val(txtYear)
    SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY' AND NOT (ED_DIV = 'BD')"
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsMain.EOF Then
        I = 0
        xNTot = rsMain.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    Do While Not rsMain.EOF
        DoEvents
        MDIMain.panHelp(0).FloodPercent = (I / xNTot) * 100: I = I + 1
        xEmpNo = rsMain("ED_EMPNBR")
        'Get glbDiv
        glbDiv = rsMain("ED_DIV")
        xSec = rsMain("ED_SECTION")

        'Get the L/LE Point for this employee - Begin
        'Get Carryover points first, if no carryover and then no reduction, because the balance can't be negative
        SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_LEPOINT FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xFDate) & " "
        SQLQ = SQLQ & "AND (AD_REASON='2100') "
        SQLQ = SQLQ & "AND AD_LEPOINT <>0 "
        rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTAtt.EOF Then
            If Not IsNull(rsTAtt("AD_LEPOINT")) Then
                If rsTAtt("AD_LEPOINT") >= 1 Then
                    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xFDate) & " "
                    SQLQ = SQLQ & "AND (AD_REASON='2400') "
                    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsTemp.EOF Then
                        rsTemp.AddNew
                        rsTemp("AD_COMPNO") = "001"
                        rsTemp("AD_EMPNBR") = xEmpNo
                        rsTemp("AD_DOA") = xFDate
                        rsTemp("AD_REASON") = "2400"
                        rsTemp("AD_HRS") = 0
                        rsTemp("AD_LEPOINT") = -1
                        rsTemp("AD_INDICATOR") = 1
                        rsTemp("AD_SEN") = 0
                        rsTemp("AD_LDATE") = Date
                        rsTemp("AD_LTIME") = Time$
                        rsTemp("AD_LUSER") = glbUserID
                        rsTemp.Update
                        'Debug.Print xEmpNo
                    Else
                        rsTemp("AD_LEPOINT") = -1
                        rsTemp("AD_INDICATOR") = 1
                        rsTemp("AD_SEN") = 0
                        rsTemp("AD_LDATE") = Date
                        rsTemp("AD_LTIME") = Time$
                        rsTemp("AD_LUSER") = glbUserID
                        rsTemp.Update
                    End If
                    rsTemp.Close
                End If
            End If
        End If
        rsTAtt.Close
        'Get the L/LE Point for this employee - End
        
        'Get the ABS Point for this employee - Begin
        'Get Carryover points first, if no carryover and then no reduction, because the balance can't be negative
        SQLQ = "SELECT AD_EMPNBR,AD_DOA,AD_POINT FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        SQLQ = SQLQ & "AND AD_DOA = " & Date_SQL(xFDate) & " "
        SQLQ = SQLQ & "AND (AD_REASON='2200') "
        SQLQ = SQLQ & "AND AD_POINT <>0 "
        rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTAtt.EOF Then
            If Not IsNull(rsTAtt("AD_POINT")) Then
                If rsTAtt("AD_POINT") >= 1 Then
                    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xFDate) & " "
                    SQLQ = SQLQ & "AND (AD_REASON='2300') "
                    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsTemp.EOF Then
                        rsTemp.AddNew
                        rsTemp("AD_COMPNO") = "001"
                        rsTemp("AD_EMPNBR") = xEmpNo
                        rsTemp("AD_DOA") = xFDate
                        rsTemp("AD_REASON") = "2300"
                        rsTemp("AD_HRS") = 0
                        rsTemp("AD_POINT") = -1
                        rsTemp("AD_INDICATOR") = 1
                        rsTemp("AD_SEN") = 0
                        rsTemp("AD_LDATE") = Date
                        rsTemp("AD_LTIME") = Time$
                        rsTemp("AD_LUSER") = glbUserID
                        rsTemp.Update
                        'Debug.Print xEmpNo
                    Else
                        rsTemp("AD_POINT") = -1
                        rsTemp("AD_INDICATOR") = 1
                        rsTemp("AD_SEN") = 0
                        rsTemp("AD_LDATE") = Date
                        rsTemp("AD_LTIME") = Time$
                        rsTemp("AD_LUSER") = glbUserID
                        rsTemp.Update
                    End If
                    rsTemp.Close
                End If
            End If
        End If
        rsTAtt.Close
        'Get the ABS Point for this employee - End
        
        rsMain.MoveNext
    Loop
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = DEFAULT
    MsgBox "Update completed"
End Function
Private Function modYTDcarryover()
Dim rsTAtt As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsMain As New ADODB.Recordset
Dim SQLQ, xNum, xCode, xSec, xEmpNo
Dim xUnexFlag, xEmlFlag, xUnexVal, xExcuVal, xEmlVal, glbDiv, glbYear, xYear, glbPointType
Dim I, xNTot, xDOA
Dim xFDate, xTDate
    xFDate = CVDate(GetMonth("Jan") & " 1," & txtYear + 1)
    xTDate = CVDate(GetMonth("Dec") & " 31," & txtYear + 1)
    xYear = Val(txtYear)
    SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY' "
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsMain.EOF Then
        I = 0
        xNTot = rsMain.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    Do While Not rsMain.EOF
        DoEvents
        MDIMain.panHelp(0).FloodPercent = (I / xNTot) * 100: I = I + 1
        xEmpNo = rsMain("ED_EMPNBR")
        'Get glbDiv
        glbDiv = rsMain("ED_DIV")
        xSec = rsMain("ED_SECTION")

        ''Get the L/LE Point for this employee - Begin
        'SQLQ = "SELECT SUM(AD_LEPOINT) AS TOTNUM FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        'SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & xYear & " "
        'SQLQ = SQLQ & "AND AD_LEPOINT <>0 "
        'rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        'If Not rsTAtt.EOF Then
        '    If Not IsNull(rsTAtt("TOTNUM")) Then
        '        If rsTAtt("TOTNUM") > 0 Then
        '            SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        '            SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xFDate) & " "
        '            SQLQ = SQLQ & "AND (AD_REASON='2100') "
        '            rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        '            If rsTemp.EOF Then
        '                rsTemp.AddNew
        '                rsTemp("AD_COMPNO") = "001"
        '                rsTemp("AD_EMPNBR") = xEmpNo
        '                rsTemp("AD_DOA") = xFDate
        '                rsTemp("AD_REASON") = "2100"
        '                rsTemp("AD_HRS") = 0
        '                rsTemp("AD_LEPOINT") = rsTAtt("TOTNUM")
        '                rsTemp("AD_INDICATOR") = 1
        '                rsTemp("AD_SEN") = 0
        '                rsTemp("AD_LDATE") = Date
        '                rsTemp("AD_LTIME") = Time$
        '                rsTemp("AD_LUSER") = glbUserID
        '                rsTemp.Update
        '                'Debug.Print xEmpNo
        '            Else
        '                If IsNull(rsTemp("AD_LEPOINT")) Then
        '                    rsTemp("AD_LEPOINT") = rsTAtt("TOTNUM")
        '                    rsTemp("AD_INDICATOR") = 1
        '                    rsTemp("AD_SEN") = 0
        '                    rsTemp("AD_LDATE") = Date
        '                    rsTemp("AD_LTIME") = Time$
        '                    rsTemp("AD_LUSER") = glbUserID
        '                    rsTemp.Update
        '                End If
        '            End If
        '            rsTemp.Close
        '        End If
        '    End If
        'End If
        'rsTAtt.Close
        ''Get the L/LE Point for this employee - End
        
        'Get the ABS Point for this employee - Begin
        SQLQ = "SELECT SUM(AD_POINT) AS TOTNUM FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & xYear & " "
        SQLQ = SQLQ & "AND AD_POINT <>0 "
        rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTAtt.EOF Then
            If Not IsNull(rsTAtt("TOTNUM")) Then
                If rsTAtt("TOTNUM") > 0 Then
                    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xFDate) & " "
                    SQLQ = SQLQ & "AND (AD_REASON='2200') "
                    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsTemp.EOF Then
                        rsTemp.AddNew
                        rsTemp("AD_COMPNO") = "001"
                        rsTemp("AD_EMPNBR") = xEmpNo
                        rsTemp("AD_DOA") = xFDate
                        rsTemp("AD_REASON") = "2200"
                        rsTemp("AD_HRS") = 0
                        rsTemp("AD_POINT") = rsTAtt("TOTNUM")
                        rsTemp("AD_INDICATOR") = 1
                        rsTemp("AD_SEN") = 0
                        rsTemp("AD_LDATE") = Date
                        rsTemp("AD_LTIME") = Time$
                        rsTemp("AD_LUSER") = glbUserID
                        rsTemp.Update
                        'Debug.Print xEmpNo
                    Else
                        If IsNull(rsTemp("AD_POINT")) Then
                            rsTemp("AD_POINT") = rsTAtt("TOTNUM")
                            rsTemp("AD_INDICATOR") = 1
                            rsTemp("AD_SEN") = 0
                            rsTemp("AD_LDATE") = Date
                            rsTemp("AD_LTIME") = Time$
                            rsTemp("AD_LUSER") = glbUserID
                            rsTemp.Update
                        End If
                    End If
                    rsTemp.Close
                End If
            End If
        End If
        rsTAtt.Close
        'Get the ABS Point for this employee - End
        
        rsMain.MoveNext
    Loop
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = DEFAULT
    MsgBox "Update completed"
End Function

Private Function modQtrEnd()
Dim rsTAtt As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim rsMain As New ADODB.Recordset
Dim SQLQ, xNum, xCode, xSec, xEmpNo
Dim xUnexFlag, xEmlFlag, xUnexVal, xExcuVal, xEmlVal, glbDiv, glbYear, xYear, glbPointType
Dim I, xNTot, xDOA
Dim xFDate, xTDate
    xFDate = CVDate(lblDRdesc(0))
    xTDate = CVDate(lblDRdesc(1))
    SQLQ = "SELECT ED_EMPNBR,ED_DIV,ED_SECTION FROM HREMP WHERE ED_SECTION='HRLY' "
    rsMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsMain.EOF Then
        I = 0
        xNTot = rsMain.RecordCount
    End If
    MDIMain.panHelp(0).FloodType = 1
    MDIMain.panHelp(1).Caption = " Please Wait"
    Do While Not rsMain.EOF
        DoEvents
        MDIMain.panHelp(0).FloodPercent = (I / xNTot) * 100: I = I + 1
        xEmpNo = rsMain("ED_EMPNBR")
        'Get glbDiv
        glbDiv = rsMain("ED_DIV")
        xSec = rsMain("ED_SECTION")
        
        '''SQLQ = "DELETE FROM HR_ATTENDANCE WHERE (1=1) "
        '''SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
        '''SQLQ = SQLQ & "AND (AD_REASON='QRED') "
        '''gdbAdoIhr001.Execute SQLQ
        '''Exit Function
        
        'Get the L/LE Point for this employee
        SQLQ = "SELECT SUM(AD_LEPOINT) AS TOTNUM FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
        SQLQ = SQLQ & "AND to_char(AD_DOA,'yyyy')=" & txtYear & " "
        SQLQ = SQLQ & "AND AD_DOA <=" & Date_SQL(xTDate) & " "
        SQLQ = SQLQ & "AND AD_LEPOINT <>0 "
        rsTAtt.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTAtt.EOF Then
            If Not IsNull(rsTAtt("TOTNUM")) Then
                If rsTAtt("TOTNUM") > 0 Then
                    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE AD_EMPNBR=" & xEmpNo & " "
                    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
                    SQLQ = SQLQ & "AND (AD_REASON='QRED') "
                    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
                    If rsTemp.EOF Then
                        rsTemp.AddNew
                        rsTemp("AD_COMPNO") = "001"
                        rsTemp("AD_EMPNBR") = xEmpNo
                        rsTemp("AD_DOA") = xTDate
                        rsTemp("AD_REASON") = "QRED"
                        rsTemp("AD_HRS") = 0
                        rsTemp("AD_LEPOINT") = -1
                        rsTemp("AD_INDICATOR") = 1
                        rsTemp("AD_SEN") = 0
                        rsTemp("AD_LDATE") = Date
                        rsTemp("AD_LTIME") = Time$
                        rsTemp("AD_LUSER") = glbUserID
                        rsTemp.Update
                        'Debug.Print xEmpNo
                    Else
                        If IsNull(rsTemp("AD_LEPOINT")) Or rsTemp("AD_LEPOINT") <> -1 Then
                            rsTemp("AD_LEPOINT") = -1
                            rsTemp("AD_INDICATOR") = 1
                            rsTemp("AD_SEN") = 0
                            rsTemp("AD_LDATE") = Date
                            rsTemp("AD_LTIME") = Time$
                            rsTemp("AD_LUSER") = glbUserID
                            rsTemp.Update
                        End If
                    End If
                    rsTemp.Close
                End If
            End If
        End If
        rsTAtt.Close
        
        rsMain.MoveNext
    Loop
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodPercent = 0
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    Screen.MousePointer = DEFAULT
    MsgBox "Update completed"
    
End Function
Public Sub cmdDelete_Click()
Dim a As Integer
'Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%, DtTm As Variant, x%
Dim DgDef, Title$, Msg$, Response%

If glbFormCaption = "Quarter End" Then
    MsgBox "This Transaction is not available"
    Exit Sub
End If

If Not chekQtrEnd() Then
  Exit Sub
End If

Title$ = "Mass Delete of " & glbFormCaption
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
Msg$ = "Are You Sure You Want To Delete ALL records for this criteria?"
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then    ' Evaluate response
    Exit Sub
End If


If modDelRecs() = True Then
  MsgBox "The records were successfully deleted"
Else
  MsgBox "You have no records in this criteria!"
End If

fglbDelete% = True

Screen.MousePointer = DEFAULT
'MsgBox "Records Deleted Successfully"
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "Attendance Entitlement", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub
Private Function modDelRecs()
Dim BD As Integer
Dim SQLQ As String, countr As Integer
Dim Dat1 As Variant, Dat2 As Variant
Dim iOneWhere As Integer, NxtSQL As String, strReas$
Dim dyn_Temp As New ADODB.Recordset
Dim xTDate
modDelRecs = False
On Error GoTo cmdDel_Err


Screen.MousePointer = HOURGLASS
xTDate = CVDate(GetMonth("Jan") & " 1," & txtYear)
If glbFormCaption = "Year End Carryover" Then
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (1=1) "
    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
    SQLQ = SQLQ & "AND (AD_REASON='2100' OR AD_REASON='2200') "
    
    dyn_Temp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adOpenStatic
    If dyn_Temp.BOF And dyn_Temp.EOF Then
      modDelRecs = False
    Else
      modDelRecs = True
    End If
    dyn_Temp.Close
    If modDelRecs Then
        SQLQ = "Delete FROM HR_ATTENDANCE WHERE (1=1) "
        SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
        SQLQ = SQLQ & "AND (AD_REASON='2100' OR AD_REASON='2200') "
        gdbAdoIhr001.Execute SQLQ
    End If
End If

If glbFormCaption = "Year End Reduction For Non BD" Then
    SQLQ = "SELECT * FROM HR_ATTENDANCE WHERE (1=1) "
    SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
    SQLQ = SQLQ & "AND (AD_REASON='2300' OR AD_REASON='2400') "
    SQLQ = SQLQ & "AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION='HRLY' AND NOT (ED_DIV='BD')) "
    
    dyn_Temp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic ', adOpenStatic
    If dyn_Temp.BOF And dyn_Temp.EOF Then
      modDelRecs = False
    Else
      modDelRecs = True
    End If
    dyn_Temp.Close
    If modDelRecs Then
        SQLQ = "Delete FROM HR_ATTENDANCE WHERE (1=1) "
        SQLQ = SQLQ & "AND AD_DOA =" & Date_SQL(xTDate) & " "
        SQLQ = SQLQ & "AND (AD_REASON='2300' OR AD_REASON='2400') "
        SQLQ = SQLQ & "AND AD_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION='HRLY' AND NOT (ED_DIV='BD')) "
        gdbAdoIhr001.Execute SQLQ
    End If
End If

Screen.MousePointer = DEFAULT


Exit Function
cmdDel_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_ATTENDANCE", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    RollBack
    Resume Next
Else
    Unload Me
End If
End Function

Private Function chekQtrEnd()
Dim dd&
Dim Msg$, DgDef As Variant, Response%
chekQtrEnd = False
On Error GoTo chkEOTHERE_Err
If Len(txtYear.Text) < 1 Then
    MsgBox "Year is a required field"
    txtYear.SetFocus
    Exit Function
End If
If Not IsNumeric(txtYear.Text) Then
    MsgBox "Invalid Year"
    txtYear.SetFocus
    Exit Function
End If
If Len(txtYear.Text) <> 4 Then
    MsgBox "Invalid Year"
    txtYear.SetFocus
    Exit Function
End If

If glbFormCaption = "Quarter End" Then
    If Len(cmdQTR.Text) < 1 Then
        MsgBox "Quarter is a required field"
        cmdQTR.SetFocus
        Exit Function
    End If
End If

chekQtrEnd = True
Exit Function

chkEOTHERE_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkCOEFlag", "HREARN", "Update")
Resume Next

End Function
Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "frmUQuarterEnd"
End Sub

Private Sub Form_Load()
glbOnTop = "frmUQuarterEnd"
Me.Caption = glbFormCaption
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
lblNewRec.Caption = ""
If glbFormCaption = "Quarter End" Then
    lblQTRNo.Visible = True
    cmdQTR.Visible = True
    cmdQTR.AddItem "1"
    cmdQTR.AddItem "2"
    cmdQTR.AddItem "3"
    cmdQTR.AddItem "4"
End If
    lblDateRange.Visible = True
    lblDRdesc(0).Visible = True
    lblDRdesc(1).Visible = True
    lblTo.Visible = True
Call INI_Controls(Me)

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmUQuarterEnd = Nothing
End Sub


Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
TF = True
UpdateState = OPENING
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False

'alpAPPNBR.Enabled = TF
End Sub
Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = True  'GetMassUpdateSecurities("Other_Earnings_MassUpdate", glbUserID)
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

Private Sub txtYear_Change()
Call SetDateRange
End Sub
