VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmImport 
   Caption         =   "Friesen Corporation"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Data1 
      Height          =   735
      Left            =   5280
      Top             =   1920
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "Import"
      Height          =   495
      Left            =   3960
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
   Begin Threed.SSPanel spShow 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   2
      Top             =   975
      Width           =   5355
      _Version        =   65536
      _ExtentX        =   9446
      _ExtentY        =   820
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderWidth     =   2
      BevelOuter      =   1
      BevelInner      =   2
      FloodType       =   1
      Alignment       =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Import Excel File"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1890
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SQLQ, xCategory, xEvent, xComments As String
Public xEventDate, xFollowUpDate
Public xEmpnbr

Private Function getRows(exSheet As Object)
Dim x
x = 1
Do While True
    If exSheet.Cells(x, 1) = "" Then
        Exit Do
    Else
        x = x + 1
    End If
Loop
getRows = x - 1
End Function
Private Function getCols(exSheet As Object)
Dim x
x = 1
Do While True
    If exSheet.Cells(1, x) = "" Then
        Exit Do
    Else
        x = x + 1
    End If
Loop
getCols = x - 1
End Function
'Function Date_SQL(xDATE) As String
'Date_SQL = " NULL "
'If IsDate(xDATE) Then
'    If glbOracle Then
'        Date_SQL = " TO_DATE('" & Format(xDATE, "DD-MM-YYYY") & "','DD-MM-YYYY') "
'    ElseIf glbSQL Then
'        Date_SQL = " ('" & Format(xDATE, "MMM DD,YYYY") & "') "
'    Else
'        Date_SQL = " CVDATE('" & xDATE & "') "
'    End If
'End If
'End Function

Private Function updFollow(xType)
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim Edit1 As Integer

updFollow = False

'On Error GoTo CrFollow_Err


    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & xEmpnbr
    SQLQ = SQLQ & " AND EF_FREAS = 'PREV'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(xEventDate)
   
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    

If xType = "U" Then
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
   
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = xEmpnbr
        rsTB("EF_FDATE") = xEventDate
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(xEmpnbr, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "PREV"
           
        rsTB("EF_COMMENTS") = xComments
              
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = "999999999"
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
       ' MsgBox Msg
        Exit Function
    End If
   
  
Exit Function

CrFollow_Err:
MsgBox (Err.Description)

End Function
Private Sub cmdImport_Click()
Dim cnReview As New ADODB.Connection
Dim exApp As Object, exBook As Object, exSheet As Object
Dim rsReview As New ADODB.Recordset
Dim xCols As Integer, xRows
Dim xCol, xRow
Dim xTitle()
Dim k
Dim FileName, ExelFName, xEmpName, xSuperName As String
Dim xRepAuth, xRepA, xlen, xPos, FileLength
'Dim iCount As Integer

Dim RSTABL As New ADODB.Recordset
Dim SearchString, SearchChar, MyPos
On Error GoTo err_update


CommonDialog1.Filter = "Exel Files|*.xls"   'open all files| *.*   '"Text File|*.txt"
CommonDialog1.ShowOpen
FileName = CommonDialog1.FileName

If FileName <> "" Then
    'Get the file name only procedure starts
   
    MyPos = InStr(FileName, "\")
    FileLength = Len(FileName)
    ExelFName = Mid(FileName, MyPos + 1, FileLength)
    Do While InStr(ExelFName, "\") <> 0
        MyPos = InStr(ExelFName, "\")
        MyPos = MyPos + 1
        ExelFName = Mid(ExelFName, MyPos, FileLength)
    'Get the file name only procedure ends
    Loop
''''''''''''''''''''''
    Label1.FontSize = 8
    Label1.Caption = "Please Wait While System Downloads(" & ExelFName & ")"
    cmdImport.Enabled = False
    Set exApp = CreateObject("Excel.Application")
    Set exBook = exApp.Workbooks.Open(FileName)
    Set exSheet = exBook.Worksheets(1)
    xRows = getRows(exSheet)
    'iCount = xRows
    If xRows > 0 Then
        k = 0
        xEmpnbr = Trim(exSheet.Cells(1, 6))
        xEmpName = Trim(exSheet.Cells(1, 1))
        xEmpnbr = CDbl(xEmpnbr)
        SQLQ = "SELECT * FROM HR_PERFORM_FRIESEN WHERE PH_EMPNBR=" & xEmpnbr
        rsReview.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
      '  If rsReview.EOF Then
            For xRow = 3 To xRows
                spShow.FloodColor = vbGreen
                spShow.FloodPercent = (xRow / xRows) * 100
                
                    rsReview.AddNew
                    rsReview("PH_COMPNO") = "001"
                    rsReview("PH_EMPNBR") = xEmpnbr
                    rsReview("PH_EMPNAME") = xEmpName
                    rsReview("PH_CURRENT") = 0
                                    
                    xEventDate = Trim(exSheet.Cells(xRow, 1))
                    If IsDate(xEventDate) Then
                        Dim dt As Date
                        dt = CDate(xEventDate)
                        If dt > Date Then
                            updFollow ("U")
                        End If
                        rsReview("PH_PREVIEW") = xEventDate
                    End If
               
                    xCategory = Trim(exSheet.Cells(xRow, 2))
                    If Len(xCategory) > 0 Then
                        If xCategory = "Attendance" Then
                            rsReview("PH_CATECODE") = "RC3"
                        
                        ElseIf xCategory = "Teamwork" Or xCategory = "Team Work" Then
                            rsReview("PH_CATECODE") = "RC4"
                        ElseIf xCategory = "Productivity" Then
                            rsReview("PH_CATECODE") = "RC1"
                        ElseIf xCategory = "Time Management" Then
                            rsReview("PH_CATECODE") = "RC2"
                        ElseIf xCategory = "Safty" Then
                            rsReview("PH_CATECODE") = "RC5"
                        End If
                    End If
                xEvent = Trim(exSheet.Cells(xRow, 3))
                If Len(xEvent) > 0 Then
                    If xEvent = "PMS Info" Then
                        rsReview("PH_EVENTCODE") = "PMS"
                    ElseIf xEvent = "Coaching" Then
                        rsReview("PH_EVENTCODE") = "COAC"
                    ElseIf xEvent = "Promotion" Then
                        rsReview("PH_EVENTCODE") = "PROM"
                       
                    ElseIf xEvent = "Review" Then
                        rsReview("PH_EVENTCODE") = "PERF"
            
                    ElseIf xEvent = "Training" Then
                        rsReview("PH_EVENTCODE") = "TR"
                    ElseIf xEvent = "PMS Rework" Then
                        rsReview("PH_EVENTCODE") = "REWK"
                    ElseIf xEvent = "PMS Skills Testing" Then
                        rsReview("PH_EVENTCODE") = "SKIL"
                    ElseIf xEvent = "PMS Update Meeting" Then
                        rsReview("PH_EVENTCODE") = "UPDT"
                    End If
                End If
                xRepAuth = Trim(exSheet.Cells(xRow, 4))
                If Len(xRepAuth) > 0 Then
                xlen = Len(xRepAuth)
                    xPos = InStr(1, xRepAuth, ":", vbTextCompare)
                    xPos = xPos - 1
                    xRepA = Left(xRepAuth, xPos)
                    xRepA = CLng(xRepA)
                    rsReview("PH_REPTAU") = xRepA
                    xSuperName = Mid(xRepAuth, xPos + 2, xlen)
                    rsReview("PH_SUPERNAME") = xSuperName
                End If
                
                xFollowUpDate = Trim(exSheet.Cells(xRow, 5))
                If IsDate(xFollowUpDate) Then
                    rsReview("PH_PNEXT") = xFollowUpDate
                End If
                xComments = Trim(exSheet.Cells(xRow, 6))
                If Len(xComments) > 0 Then
                    rsReview("PH_COMMENTS") = xComments
                End If
        
                rsReview("PH_LDATE") = Date
                rsReview("PH_LTIME") = Time$
                rsReview("PH_LUSER") = "999999999"
                rsReview.Update
               
            Next
            rsReview.Close
            
            ''modify current
            SQLQ = "SELECT * FROM HR_PERFORM_FRIESEN WHERE PH_EMPNBR=" & xEmpnbr
            SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC"
            rsReview.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
             
            rsReview("PH_CURRENT") = 1
            rsReview.Update
            rsReview.Close
            '''
            
'       ' Else
'            MsgBox "The (" & ExelFName & ") File already imported", vbCritical, "info:HR"
'            Set exSheet = Nothing
'            Set exBook = Nothing
'            exApp.Quit
'            Set exApp = Nothing
'            Screen.MousePointer = vbDefault
'            cmdImport.Enabled = True
'            Label1.Caption = "Import An Other Excel File"
'            Exit Sub
'        End If
            
        
        'STARTS''''''''''''' ADDING CODE FOR "PERFORMANCE CATEGORY REVIEW" IN HRTABL
                    '''''ADDING CODE BY DEFAULT FOR Productivity'''''''''
'                    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPG' AND TB_KEY = 'RC1' "
'                    RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    If RSTABL.EOF Then
'                       RSTABL.AddNew
'                       RSTABL("TB_COMPNO") = "001"
'                       RSTABL("TB_NAME") = "SDPG"
'                       RSTABL("TB_KEY") = "RC1"
'                       RSTABL("TB_DESC") = "Productivity"
'                       RSTABL("TB_LDATE") = Date
'                       RSTABL("TB_LTIME") = Time$
'                       RSTABL("TB_LUSER") = "999999999"
'                       RSTABL.Update
'                     End If
'                     RSTABL.Close
'                    '''''''''''''
'                    'ADDING CODE BY DEFAULT FOR Time Management
'                    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPG' AND TB_KEY = 'RC2' "
'                    RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPG"
'                            RSTABL("TB_KEY") = "RC2"
'                            RSTABL("TB_DESC") = "Time Management"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'                    '''''''''''''
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPG' AND TB_KEY = 'RC3' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPG"
'                            RSTABL("TB_KEY") = "RC3"
'                            RSTABL("TB_DESC") = "Attendance"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'                    '''''''''
'                            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPG' AND TB_KEY = 'RC4' "
'                            RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                            If RSTABL.EOF Then
'                                RSTABL.AddNew
'                                RSTABL("TB_COMPNO") = "001"
'                                RSTABL("TB_NAME") = "SDPG"
'                                RSTABL("TB_KEY") = "RC4"
'                                RSTABL("TB_DESC") = "Team Work"
'                                RSTABL("TB_LDATE") = Date
'                                RSTABL("TB_LTIME") = Time$
'                                RSTABL("TB_LUSER") = "999999999"
'                                RSTABL.Update
'                            End If
'                            RSTABL.Close
'                    '''''''
'                    'Adding code for Safty by default
'                            SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPG' AND TB_KEY = 'RC5' "
'                            RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                            If RSTABL.EOF Then
'                                RSTABL.AddNew
'                                RSTABL("TB_COMPNO") = "001"
'                                RSTABL("TB_NAME") = "SDPG"
'                                RSTABL("TB_KEY") = "RC5"
'                                RSTABL("TB_DESC") = "Safety"
'                                RSTABL("TB_LDATE") = Date
'                                RSTABL("TB_LTIME") = Time$
'                                RSTABL("TB_LUSER") = "999999999"
'                                RSTABL.Update
'                            End If
'                            RSTABL.Close
'                'ENDS''''''''''''' ADDING CODE FOR "PERFORMANCE CATEGORY REVIEW" IN HRTABL
'
'            'STARTS''''''''''''' ADDING CODE FOR "PERFORMANCE EVENT REVIEW" IN HRTABL
'                    ''''''Adding Event Code by default for 1 Reward - $25 Gift Certificate
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'RW25' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "RW25"
'                            RSTABL("TB_DESC") = "1 Reward-$25 Gift Certificate"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            ''''''
'            ''''''Adding Event Code by default for 2 Reward - Sports Tickets
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'RWST' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "RWST"
'                            RSTABL("TB_DESC") = "2 Reward - Sports Tickets"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            ''''''
'            ''''''Adding Event Code by default for 1 Reward - 3 Reward - Movie Tickets
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'RWMT' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "RWMT"
'                            RSTABL("TB_DESC") = "3 Reward - Movie Tickets"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            ''''''
'            ''''''Adding Event Code by default for 4 Reward - Lunch Thank-You Card
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'RWLU' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "RWLU"
'                            RSTABL("TB_DESC") = "4 Reward-Lunch Thank-You Card"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            ''''''
'            ''''''Adding Event Code by default for 5 Reward - Thank-You Card
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'RWTY' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "RWTY"
'                            RSTABL("TB_DESC") = "5 Reward - Thank-You Card"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            ''''''
'            ''''''Adding Event Code by default for 6 Reward - I Want You to Know...
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'RWW' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "RWW"
'                            RSTABL("TB_DESC") = "6 Reward-I Want You to Know..."
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            ''''''
'                    SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'PMS' "
'                    RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                    If RSTABL.EOF Then
'                        RSTABL.AddNew
'                        RSTABL("TB_COMPNO") = "001"
'                        RSTABL("TB_NAME") = "SDPE"
'                        RSTABL("TB_KEY") = "PMS"
'                        RSTABL("TB_DESC") = "PMS Info"
'                        RSTABL("TB_LDATE") = Date
'                        RSTABL("TB_LTIME") = Time$
'                        RSTABL("TB_LUSER") = "999999999"
'                        RSTABL.Update
'                    End If
'                    RSTABL.Close
'            '''
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'COAC' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "COAC"
'                            RSTABL("TB_DESC") = "PMS Coaching"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            '''
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'PROM' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "PROM"
'                            RSTABL("TB_DESC") = "PMS Promotion"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'
'            '''
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'PERF' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "PERF"
'                            RSTABL("TB_DESC") = "Performance Review"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            '''
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'TR' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "TR"
'                            RSTABL("TB_DESC") = "PMS Training"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'            '''
'
'               'Adding code for PMS Rework
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'SKIL' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "SKIL"
'                            RSTABL("TB_DESC") = "PMS Skills Testing"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'                ''
'                 'Adding code for PMS Skills Testing
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'SKIL' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "SKIL"
'                            RSTABL("TB_DESC") = "PMS Skills Testing"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
'                ''
'                'Adding code for PMS  Update Meeting
'                        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = 'SDPE' AND TB_KEY = 'UPDT' "
'                        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
'                        If RSTABL.EOF Then
'                            RSTABL.AddNew
'                            RSTABL("TB_COMPNO") = "001"
'                            RSTABL("TB_NAME") = "SDPE"
'                            RSTABL("TB_KEY") = "UPDT"
'                            RSTABL("TB_DESC") = "PMS Update Meeting"
'                            RSTABL("TB_LDATE") = Date
'                            RSTABL("TB_LTIME") = Time$
'                            RSTABL("TB_LUSER") = "999999999"
'                            RSTABL.Update
'                        End If
'                        RSTABL.Close
                ''
            
    
    'ENDS''''''''''''' ADDING CODE FOR "PERFORMANCE EVENT REVIEW" IN HRTABL

           
    
    
        Set exSheet = Nothing
        Set exBook = Nothing
        exApp.Quit
        Set exApp = Nothing
        Screen.MousePointer = vbDefault
        spShow.FloodPercent = 100
        MsgBox "Imported (" & ExelFName & ") File Successfully", , "info:HR"
        
        cmdImport.Enabled = True
        Label1.Caption = "Import An Other Excel File"
        'End
    End If
   
End If
Exit Sub
err_update:
 Me.Hide
 MsgBox Err.Description
 End
End Sub

Private Sub Form_Load()
'Call modLoadINI
'gdbAdoIhr001.CommandTimeout = 300
'gdbAdoIhr001.Open glbAdoIHRDB
'gdbAdoIhr001X.Open Replace(glbAdoIHRDB, "IHR001", "ihr001x")

End Sub
