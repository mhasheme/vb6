Attribute VB_Name = "MCNSTNTS"
Option Explicit

' Base declarations from constants.txt - MS basic requirements excerpts only
'=================================================
'ZOrder Method
Global Const BRINGTOFRONT = 0
Global Const SENDTOBACK = 1
' Event Parameters

' Properties

' Colors
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF
Global Const GREY = &HC0C0C0
Global Const DGREY = &HC0C0C0

' Align (picture box)
Global Const NONE = 0
Global Const ALIGN_TOP = 1
Global Const ALIGN_BOTTOM = 2

' Alignment
Global Const LEFT_JUSTIFY = 0  ' 0 - Left Justify
Global Const RIGHT_JUSTIFY = 1 ' 1 - Right Justify
Global Const CENTER = 2        ' 2 - Center

' BorderStyle (form)
'Global Const NONE = 0          ' 0 - None
Global Const FIXED_SINGLE = 1   ' 1 - Fixed Single
Global Const SIZABLE = 2        ' 2 - Sizable (Forms only)
Global Const FIXED_DOUBLE = 3   ' 3 - Fixed Double (Forms only)

              
' MousePointer
Global Const DEFAULT = 0        ' 0 - Default
Global Const ARROW = 1          ' 1 - Arrow
Global Const CROSSHAIR = 2      ' 2 - Cross
Global Const IBEAM = 3          ' 3 - I-Beam
Global Const ICON_POINTER = 4   ' 4 - Icon
Global Const SIZE_POINTER = 5   ' 5 - Size
Global Const SIZE_NE_SW = 6     ' 6 - Size NE SW
Global Const SIZE_N_S = 7       ' 7 - Size N S
Global Const SIZE_NW_SE = 8     ' 8 - Size NW SE
Global Const SIZE_W_E = 9       ' 9 - Size W E
Global Const UP_ARROW = 10      ' 10 - Up Arrow
Global Const HOURGLASS = 11     ' 11 - Hourglass
Global Const NO_DROP = 12       ' 12 - No drop

' DragMode
Global Const MANUAL = 0    ' 0 - Manual
Global Const AUTOMATIC = 1 ' 1 - Automatic

' WindowState
Global Const NORMAL = 0    ' 0 - Normal
Global Const MINIMIZED = 1 ' 1 - Minimized
Global Const MAXIMIZED = 2 ' 2 - Maximized

' MsgBox parameters
Global Const MB_OK = 0                 ' OK button only
Global Const MB_OKCANCEL = 1           ' OK and Cancel buttons
Global Const MB_ABORTRETRYIGNORE = 2   ' Abort, Retry, and Ignore buttons
Global Const MB_YESNOCANCEL = 3        ' Yes, No, and Cancel buttons
Global Const MB_YESNO = 4              ' Yes and No buttons
Global Const MB_RETRYCANCEL = 5        ' Retry and Cancel buttons

Global Const MB_ICONSTOP = 16          ' Critical message
Global Const MB_ICONQUESTION = 32      ' Warning query
Global Const MB_ICONEXCLAMATION = 48   ' Warning message
Global Const MB_ICONINFORMATION = 64   ' Information message

Global Const MB_APPLMODAL = 0          ' Application Modal Message Box
Global Const MB_DEFBUTTON1 = 0         ' First button is default
Global Const MB_DEFBUTTON2 = 256       ' Second button is default
Global Const MB_DEFBUTTON3 = 512       ' Third button is default
Global Const MB_SYSTEMMODAL = 4096      'System Modal

' MsgBox return values
Global Const IDOK = 1                  ' OK button pressed
Global Const IDCANCEL = 2              ' Cancel button pressed
Global Const IDABORT = 3               ' Abort button pressed
Global Const IDRETRY = 4               ' Retry button pressed
Global Const IDIGNORE = 5              ' Ignore button pressed
Global Const IDYES = 6                 ' Yes button pressed
Global Const IDNO = 7                  ' No button pressed

Global glbMsgCustomVal As Integer      'Samuel Ticket #20371 Franks 05/25/2011

' SetAttr, Dir, GetAttr functions
Global Const ATTR_NORMAL = 0
Global Const ATTR_READONLY = 1
Global Const ATTR_HIDDEN = 2
Global Const ATTR_SYSTEM = 4
Global Const ATTR_VOLUME = 8
Global Const ATTR_DIRECTORY = 16
Global Const ATTR_ARCHIVE = 32

'Grid
'ColAlignment,FixedAlignment Properties
Global Const GRID_ALIGNLEFT = 0
Global Const GRID_ALIGNRIGHT = 1
Global Const GRID_ALIGNCENTER = 2

'Fillstyle Property
Global Const GRID_SINGLE = 0
Global Const GRID_REPEAT = 1


'Data control
'Error event Response arguments
Global Const DATA_ERRCONTINUE = 0
Global Const DATA_ERRDISPLAY = 1

'Editmode property values
Global Const DATA_EDITNONE = 0
Global Const DATA_EDITMODE = 1
Global Const DATA_EDITADD = 2

' Options property values
Global Const DATA_DENYWRITE = &H1
Global Const DATA_DENYREAD = &H2
Global Const DATA_READONLY = &H4
Global Const DATA_APPENDONLY = &H8
Global Const DATA_INCONSISTENT = &H10
Global Const DATA_CONSISTENT = &H20
Global Const DATA_SQLPASSTHROUGH = &H40

'Validate event Action arguments
Global Const DATA_ACTIONCANCEL = 0
Global Const DATA_ACTIONMOVEFIRST = 1
Global Const DATA_ACTIONMOVEPREVIOUS = 2
Global Const DATA_ACTIONMOVENEXT = 3
Global Const DATA_ACTIONMOVELAST = 4
Global Const DATA_ACTIONADDNEW = 5
Global Const DATA_ACTIONUPDATE = 6
Global Const DATA_ACTIONDELETE = 7
Global Const DATA_ACTIONFIND = 8
Global Const DATA_ACTIONBOOKMARK = 9
Global Const DATA_ACTIONCLOSE = 10
Global Const DATA_ACTIONUNLOAD = 11



'Common Dialog Control
'Action Property
Global Const DLG_FILE_OPEN = 1
Global Const DLG_FILE_SAVE = 2
Global Const DLG_COLOR = 3
Global Const DLG_FONT = 4
Global Const DLG_PRINT = 5
Global Const DLG_HELP = 6

'File Open/Save Dialog Flags
Global Const OFN_READONLY = &H1&
Global Const OFN_OVERWRITEPROMPT = &H2&
Global Const OFN_HIDEREADONLY = &H4&
Global Const OFN_NOCHANGEDIR = &H8&
Global Const OFN_SHOWHELP = &H10&
Global Const OFN_NOVALIDATE = &H100&
Global Const OFN_ALLOWMULTISELECT = &H200&
Global Const OFN_EXTENSIONDIFFERENT = &H400&
Global Const OFN_PATHMUSTEXIST = &H800&
Global Const OFN_FILEMUSTEXIST = &H1000&
Global Const OFN_CREATEPROMPT = &H2000&
Global Const OFN_SHAREAWARE = &H4000&
Global Const OFN_NOREADONLYRETURN = &H8000&


'Printer Dialog Flags
Global Const PD_ALLPAGES = &H0&
Global Const PD_SELECTION = &H1&
Global Const PD_PAGENUMS = &H2&
Global Const PD_NOSELECTION = &H4&
Global Const PD_NOPAGENUMS = &H8&
Global Const PD_COLLATE = &H10&
Global Const PD_PRINTTOFILE = &H20&
Global Const PD_PRINTSETUP = &H40&
Global Const PD_NOWARNING = &H80&
Global Const PD_RETURNDC = &H100&
Global Const PD_RETURNIC = &H200&
Global Const PD_RETURNDEFAULT = &H400&
Global Const PD_SHOWHELP = &H800&
Global Const PD_USEglbDevModeCOPIES = &H40000
Global Const PD_DISABLEPRINTTOFILE = &H80000
Global Const PD_HIDEPRINTTOFILE = &H100000

'Help Constants
Global Const HELP_CONTEXT = &H1           'Display topic in ulTopic
Global Const HELP_QUIT = &H2              'Terminate help
Global Const HELP_INDEX = &H3             'Display index
Global Const HELP_CONTENTS = &H3
Global Const HELP_HELPONHELP = &H4        'Display help on using help
Global Const HELP_SETINDEX = &H5          'Set the current Index for multi index help
Global Const HELP_SETCONTENTS = &H5
Global Const HELP_CONTEXTPOPUP = &H8
Global Const HELP_FORCEFILE = &H9
Global Const HELP_KEY = &H101             'Display topic for keyword in offabData
Global Const HELP_COMMAND = &H102
Global Const HELP_PARTIALKEY = &H105      'call the search engine in winhelp

'Error Constants
Global Const CDERR_DIALOGFAILURE = -32768

Global Const CDERR_GENERALCODES = &H7FFF
Global Const CDERR_STRUCTSIZE = &H7FFE
Global Const CDERR_INITIALIZATION = &H7FFD
Global Const CDERR_NOTEMPLATE = &H7FFC
Global Const CDERR_NOHINSTANCE = &H7FFB
Global Const CDERR_LOADSTRFAILURE = &H7FFA
Global Const CDERR_FINDRESFAILURE = &H7FF9
Global Const CDERR_LOADRESFAILURE = &H7FF8
Global Const CDERR_LOCKRESFAILURE = &H7FF7
Global Const CDERR_MEMALLOCFAILURE = &H7FF6
Global Const CDERR_MEMLOCKFAILURE = &H7FF5
Global Const CDERR_NOHOOK = &H7FF4

'Added for CMDIALOG.VBX
Global Const CDERR_CANCEL = &H7FF3
Global Const CDERR_NODLL = &H7FF2
Global Const CDERR_ERRPROC = &H7FF1
Global Const CDERR_ALLOC = &H7FF0
Global Const CDERR_HELP = &H7FEF

'Grids
Global Const G_HORIZONTAL = 1
Global Const G_VERTICAL = 2

'---------------------------------------
'Spin Button
'---------------------------------------
'SpinOrientation
Global Const SPIN_VERTICAL = 0
Global Const SPIN_HORIZONTAL = 1


'---------------------------------------
'Masked Edit Control
'---------------------------------------
'ClipMode
Global Const ME_INCLIT = 0
Global Const ME_EXCLIT = 1


'=================================================
'Global constants - for About popup
Global Const GFSR_SystemResources% = 0
Global Const GFSR_GDIResources% = 1
Global Const GFSR_UserResources% = 2
Global Const GFSR_NextResources% = 3

'Ticket #23386 Franks 03/22/2013 - begin
Dim rsEntMain As New ADODB.Recordset
Dim xmedLTServ(24)
Dim xmedGTServ(24)
Dim xmedEntitle(24)
Dim xmedMax(24)
Dim xmedVacation(24)
Dim xoptD(24)
Dim xoptH(24)
Dim xoptF(24)
Dim xFirstMonEnt 'Ticket #23385 Franks 03/25/2013
Dim isConYear As Boolean 'Ticket #23385 Franks 03/25/2013 - Is it the conversion year
'Ticket #23386 Franks 03/22/2013 - end

'
'-------------------------------------------------------
' Centers the passed form just above center on the screen
'-------------------------------------------------------
Sub CenterForm(x As Form)
  
    Screen.MousePointer = HOURGLASS
    x.Top = (Screen.Height * 0.95) / 2 - x.Height / 2
    x.Left = Screen.Width / 2 - x.Width / 2
    x.Show
    Screen.MousePointer = DEFAULT

End Sub

Function gdbIhr001_Opn() As Integer
On Error GoTo gdbIhr001_Err

' ##########
' dkostka - 12/13/2001 - Changed to add locking parameters, this way JET will try LOCK_RETRYS number
'   of times waiting LOCK_DELAY millisecs between each try before giving up, instead of just failing
'   with 'couldn't update, currently locked' right away.
' ##########

' dkostka - 12/13/2001 - See above comment.
Const LOCK_RETRYS = 25
Const LOCK_DELAY = 200
gdbAdoIhr001.CursorLocation = adUseClient
gdbAdoIhr001X.CursorLocation = adUseClient
gdbAdoIhr001W.CursorLocation = adUseClient

gdbIhr001_Opn = False
If gdbAdoIhr001.State = adStateOpen Then gdbAdoIhr001.Close
If gdbAdoIhr001W.State = adStateOpen Then gdbAdoIhr001W.Close
If gdbAdoIhr001X.State = adStateOpen Then gdbAdoIhr001X.Close
'gdbAdoIhr001.CommandTimeout = 600
'gdbAdoIhr001X.CommandTimeout = 600
'gdbAdoIhr001W.CommandTimeout = 600
'Ticket #27856 Franks 12/03/2015 - make timeout to 600 from 300
gdbAdoIhr001.CommandTimeout = 600
gdbAdoIhr001X.CommandTimeout = 600
gdbAdoIhr001W.CommandTimeout = 600

gdbAdoIhr001.Mode = adModeReadWrite
gdbAdoIhr001.Open glbAdoIHRDB


' dkostka - 12/13/2001 - See above comment.
' sam add for test
If Not glbOracle And Not glbSQL Then
    gdbAdoIhr001.Properties("Jet OLEDB:Lock Retry") = LOCK_RETRYS
    gdbAdoIhr001.Properties("Jet OLEDB:Lock Delay") = LOCK_DELAY
End If

gdbAdoIhr001W.Mode = adModeReadWrite
gdbAdoIhr001W.Open glbAdoIHRDBW
' dkostka - 12/13/2001 - See above comment.
If Not glbOracle And Not glbSQL Then
    gdbAdoIhr001W.Properties("Jet OLEDB:Lock Retry") = LOCK_RETRYS
    gdbAdoIhr001W.Properties("Jet OLEDB:Lock Delay") = LOCK_DELAY
End If

gdbAdoIhr001X.Mode = adModeReadWrite
gdbAdoIhr001X.Open glbAdoIHRAUDIT
' dkostka - 12/13/2001 - See above comment.
If Not glbOracle And Not glbSQL Then
    gdbAdoIhr001X.Properties("Jet OLEDB:Lock Retry") = LOCK_RETRYS
    gdbAdoIhr001X.Properties("Jet OLEDB:Lock Delay") = LOCK_DELAY
End If

If isEDU Then
If gdbAdoIHREDU.State = adStateOpen Then gdbAdoIHREDU.Close
gdbAdoIHREDU.Mode = adModeReadWrite
gdbAdoIHREDU.Open glbadoIHREDU
If Not glbOracle And Not glbSQL Then
    gdbAdoIHREDU.Properties("Jet OLEDB:Lock Retry") = LOCK_RETRYS
    gdbAdoIHREDU.Properties("Jet OLEDB:Lock Delay") = LOCK_DELAY
    End If
End If
If Not glbOracle And Not glbSQL Then
    If Left(glbDBDir, 8) <> "00000000" Then
        If Dir(glbDBDir & "IHREDU.MDB") <> "" Then
            Call LinkTables(glbDBDir)
        End If
    End If
End If
'------start     'Jaddy 8/9/99
glbWFC = False
If Not glbSQL And Not glbOracle Then
    If File(glbIHRWFC) Then
        glbWFC = True
        If gdbAdoIhrWFC.State = adStateOpen Then gdbAdoIhrWFC.Close
        gdbAdoIhrWFC.Mode = adModeReadWrite
        gdbAdoIhrWFC.Open glbAdoIHRWFC
    End If
    If Left(glbIHRDBO, 8) <> "00000000" Then
        If Dir$(glbIHRDBO) <> "" Then
            If gdbAdoIhr001O.State = adStateOpen Then gdbAdoIhr001O.Close
            gdbAdoIhr001O.Mode = adModeReadWrite
            gdbAdoIhr001O.Open glbAdoIHRDBO
        End If
    End If
    If glbCBrant Then
        If Left(glbIHRDBB, 8) <> "00000000" Then
            If Dir$(glbIHRDBB) <> "" Then
                If gdbAdoIhr001B.State = adStateOpen Then gdbAdoIhr001B.Close
                gdbAdoIhr001B.Mode = adModeReadWrite
                gdbAdoIhr001B.Open glbAdoIHRDBB
            End If
        End If
    End If
    If glbGuelph Then
        If Dir$(glbSN2322) <> "" Then
            If gdbAdoSN2322.State = adStateOpen Then gdbAdoSN2322.Close
            gdbAdoSN2322.Mode = adModeReadWrite
            gdbAdoSN2322.Open glbAdoSN2322
        End If
    End If
Else
    'Ticket #24677 - SQL Conversion
    If glbGuelph Then
        'If Dir$(glbSN2322) <> "" Then
            If gdbAdoSN2322.State = adStateOpen Then gdbAdoSN2322.Close
            gdbAdoSN2322.Mode = adModeReadWrite
            'Ticket #24677 - SQL Conversion
            'gdbAdoSN2322.Open glbAdoSN2322
            gdbAdoSN2322.Open glbAdoIHRDB
        'End If
    End If
End If


'Call INICodes
glbVadim = isVadim
glbAdv = isAdvancedTracker
glbGP = isGreatPlains 'geo added
glbMediPay = isMediPay 'geo added
glbInsync = isInsync And Not isVadim
glbPayWeb = isIHRPWeb 'Ticket #20209 Franks 04/21/2011


gdbIhr001_Opn = True

Exit Function

gdbIhr001_Err:
Select Case Err
    Case 53
        MsgBox "Cannot find file, please select another."
    Case 68
        MsgBox "Drive is not responding "
    Case 55
        MsgBox "File is open - you can not perform exclusive lock'"
    Case 70
        MsgBox "Permission denied - some one is using it EXCLUSIVELY"
    Case Else
        ' dkostka - 03/21/01 - Removed useless error message, we call the global error handler after anyways.
        'MsgBox "Error occured # " & Err & " on frmSECURITY open - call support"
End Select

If Err.Number = -2147467259 Or Err.Number = -2147217843 Then
    MsgBox "Check which database you are trying to access.", vbCritical, "Invalid Database Connection"
    Exit Function
End If

glbFrmCaption$ = "gdbIhr001_Opn"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Get_Code", "TABL", "SELECT")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
End If

Exit Function
Resume 'for testing only...
End Function

Function Just_FName(FName)

Dim lnam As Integer, countr  As Integer
   
    Just_FName = FName
    lnam = Len(FName)
    countr = lnam
    Do While countr > 0
        If Mid$(FName, countr, 1) = "\" Then
            If countr = 1 Then
                countr = 2 ' so doesn't assume start 0
            End If
            Just_FName = Right$(FName, lnam - countr)
            Exit Do
        Else
            countr = countr - 1
        End If
    Loop

End Function

Sub MenuAbout()
Dim Msg As String
Dim Version As String
Dim EXEDate As Date

If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHR.EXE") <> "" Then
    EXEDate = FileDateTime(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHR.EXE")
ElseIf Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRDEMO.EXE") <> "" Then
    EXEDate = FileDateTime(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRDEMO.EXE")
End If

Version = App.Major & "." & App.Minor & "." & App.Revision
Msg = "             HR Systems Strategies Inc."
Msg = Msg & Chr(10) & "                        Toronto, ON      " & Chr(10)
Msg = Msg & Chr(10) & "             Copyright " & Chr(169) & " info:HR 2006"
Msg = Msg & Chr(10) & "                  All Rights Reserved." & Chr(10)
Msg = Msg & Chr(10) & "       Windows Version " & Version & "  " & Format(EXEDate, "Long Date")
If glbOracle Then
    Msg = Msg & Chr(10) & "           ORACLE Server 9i Product"
ElseIf glbSQL Then
    Msg = Msg & Chr(10) & "        MS SQL Server 7.0/2000 Product"
Else
    Msg = Msg & Chr(10) & "       MS Access 2002 Database Product"
End If
Msg = Msg & Chr(10) & "           Crystal Reports Version 8.0" & Chr(10)
Msg = Msg & Chr(10) & "           For support call  (416) 599-4747 x2022"
Msg = Msg & Chr(10) & "                                     1-800-567-4254" & Chr(10)
Msg = Msg & Chr(10) & "           E-Mail   support@infohr.com" & Chr(10)
Msg = Msg & Chr(10) & "           Fax         (416) 599-5031" & Chr(10)
Msg = Msg & Chr(10) & "       General Information:  www.infohr.com      "

Msg = Msg & Chr(10)

MsgBox Msg, , "info:HR"

End Sub


Sub mnu_Exit()
Dim Msg As String, DgDef As Variant, Response As Integer

Msg = "Exit info:HR?"
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
Response = MsgBox(Msg, DgDef, "Exit")
If Response = IDYES Then
    End
End If

End Sub

Sub SysInfo()
Dim Msg As String
Dim FileName As String
Dim xPath
If glbSQL Then
    Msg = Msg & Chr(10) & "SQL SERVER:  " & vbTab & " - " & SQLServerName
    Msg = Msg & Chr(10) & "DATABASE: " & vbTab & " - " & SQLDatabaseName
ElseIf glbOracle Then
    Msg = Msg & Chr(10) & "SERVER:  " & vbTab & " - " & SQLServerName
Else
    Msg = Msg & Chr(10) & "IHR001: " & vbTab & vbTab & " - " & glbIHRDB
    Msg = Msg & Chr(10) & "IHR001X: " & vbTab & " - " & glbIHRAUDIT
    If Left(glbIHRWFC, 8) <> "00000000" Then Msg = Msg & Chr(10) & "IHRWFC: " & vbTab & " - " & glbIHRWFC
    If Left(glbIHRDBO, 8) <> "00000000" Then Msg = Msg & Chr(10) & "IHROPUS: " & vbTab & " - " & glbIHRDBO
End If

Msg = Msg & Chr(10) & "REPORTS: " & vbTab & " - " & glbIHRREPORTS
Msg = Msg & Chr(10) & "EXEs: " & vbTab & vbTab & " - " & App.Path
Msg = Msg & Chr(10) & "Language: " & vbTab & " - " & IIf(glbFrench, "French", "English")
Msg = Msg & Chr(10)

' dkostka - Added file time/date information
' Get the first filename
xPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
FileName = Dir(xPath & "*.EXE")
'glbIHRREPORTS & IIf(Right(glbIHRREPORTS, 1) = "\", "", "\") & "*.EXE")
' dkostka - 04/05/2000 - Added check, so that if filename is blank (no EXE files in reports directory), IHR won't crash.
If FileName <> "" Then
    Do
        Msg = Msg & Chr(10) & UCase(FileName) & IIf(Len(FileName) < 8, vbTab & vbTab, vbTab) & "- " & FileDateTime(xPath & FileName)
        ' Get the next filename
        FileName = Dir()
    Loop While FileName <> ""
End If

' End changes by dkostka

MsgBox Msg, , "info:HR - System Information"

End Sub

Function File(FileName)
Dim x
File = True
On Error GoTo errfalse
x = GetAttr(FileName)
Exit Function
errfalse:
  If Err = 53 Then File = False
  Resume Next
End Function

Sub INICodes()
Dim x, xName
On Error GoTo INI_Err
Dim rsTABL As New ADODB.Recordset
Dim CodeDes As New ADODB.Recordset
CodeDes.Open "HRTABDES", gdbAdoIhr001, adOpenStatic
ReDim glbTblDesc(CodeDes.RecordCount + 3)
glbTblName.Add glbTblDesc(1), "DEPT"
glbTblDesc(1).Add "Departments", "HRTITLE"
rsTABL.Open "HRDEPT", gdbAdoIhr001, adOpenStatic
Do Until rsTABL.EOF
    glbTblDesc(1).Add CStr(rsTABL("DF_NAME")), CStr(rsTABL("DF_NBR"))
    rsTABL.MoveNext
Loop
rsTABL.Close

glbTblName.Add glbTblDesc(2), "DIV"
glbTblDesc(2).Add "Divisions", "HRTITLE"
rsTABL.Open "HR_DIVISION", gdbAdoIhr001, adOpenStatic
Do Until rsTABL.EOF
    glbTblDesc(2).Add CStr(rsTABL("Division_Name")), CStr(rsTABL("DIV"))
    rsTABL.MoveNext
Loop
rsTABL.Close

glbTblName.Add glbTblDesc(3), "EENAME"
glbTblDesc(3).Add "Find Employees", "HRTITLE"
rsTABL.Open "SELECT ED_EMPNBR,ED_FNAME,ED_SURNAME FROM HREMP", gdbAdoIhr001, adOpenStatic
Do Until rsTABL.EOF
    If Not IsNull(rsTABL("ED_SURNAME")) And Not IsNull(rsTABL("ED_FNAME")) Then
        xName = rsTABL("ED_SURNAME") & ", " & rsTABL("ED_FNAME")
    Else
        xName = ""
    End If
    glbTblDesc(3).Add xName, CStr(rsTABL("ED_EMPNBR"))
    rsTABL.MoveNext
Loop
rsTABL.Close
x = 4
Do Until CodeDes.EOF
    glbTblName.Add glbTblDesc(x), CodeDes("TD_NAME")
    glbTblDesc(x).Add CodeDes("TD_DESC"), "HRTITLE"
    rsTABL.Open "SELECT TB_KEY,TB_DESC FROM HRTABL WHERE TB_NAME='" & CodeDes("TD_NAME") & "'", gdbAdoIhr001, adOpenStatic
    Do Until rsTABL.EOF
        If Not IsNull(rsTABL("TB_KEY")) Then
            If Not IsNull(rsTABL("TB_DESC")) Then
                glbTblDesc(x).Add CStr(rsTABL("TB_DESC")), CStr(rsTABL("TB_KEY"))
            Else
                glbTblDesc(x).Add "", CStr(rsTABL("TB_KEY"))
            End If
        End If
        rsTABL.MoveNext
    Loop
    rsTABL.Close
    x = x + 1
    CodeDes.MoveNext
Loop
Exit Sub

INI_Err:
Select Case Err
    Case 457
    Screen.MousePointer = DEFAULT
    Resume Next
End Select


End Sub


   
   
Sub LinkTables(xDBDir As String)
Dim adoCat As New ADOX.Catalog
Dim adoTbl() As New ADOX.Table
Dim xTables As New Collection
Dim x
On Error GoTo err_LinkTables
Set adoCat = New ADOX.Catalog
Set adoCat.ActiveConnection = gdbAdoIhr001W


xTables.Add "HRTABL"
xTables.Add "HRPARCO"

xTables.Add "HREMP"
xTables.Add "HRJOB"
xTables.Add "HR_JOB_COURSE"
xTables.Add "HR_DIVISION"
xTables.Add "HR_SECURE_BASIC"
ReDim adoTbl(xTables.count)

For x = 1 To xTables.count
    Set adoTbl(x).ParentCatalog = adoCat
    If isTableExist(xTables(x), adoCat) Then
        Set adoTbl(x) = adoCat.Tables(xTables(x))
        If UCase(adoTbl(x).Properties("Jet OLEDB:Link Datasource")) <> UCase(xDBDir & "IHR001.mdb") Then
            adoTbl(x).Properties("Jet OLEDB:Link Datasource") = xDBDir & "IHR001.mdb"
            adoTbl(x).Properties("Jet OLEDB:Link Provider String") = "MS Access;Pwd=petman"
        End If
    Else
        adoTbl(x).name = xTables(x)
        adoTbl(x).Properties("Jet OLEDB:Link Datasource") = xDBDir & "IHR001.mdb"
        adoTbl(x).Properties("Jet OLEDB:Link Provider String") = "MS Access;Pwd=petman"
        adoTbl(x).Properties("Jet OLEDB:Remote Table Name") = xTables(x)
        adoTbl(x).Properties("Jet OLEDB:Create Link") = True
        adoCat.Tables.Append adoTbl(x)
        adoCat.Tables.Refresh
    End If
Next
Exit Sub
err_LinkTables:
MsgBox Err.Description
Resume Next
End Sub

Function isTableExist(TableName, adoCat As ADOX.Catalog) As Boolean
Dim adoTbl As New ADOX.Table
isTableExist = False
For Each adoTbl In adoCat.Tables
   If UCase(adoTbl.name) = UCase(TableName) Then
        isTableExist = True
   End If
Next
End Function

Function isSamuelPenCal() As Boolean 'SamuelPenCal.dat
On Error GoTo notPenCal


    If File(App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "SamuelPenCal.dat") Then
        isSamuelPenCal = True
    Else
        isSamuelPenCal = False
    End If

Exit Function

notPenCal:
    isSamuelPenCal = False
End Function
Function isInsync() As Boolean
Dim rsInsync As New ADODB.Recordset

On Error GoTo notInsync

'If glbSQL Then
If glbSQL Or glbOracle Then
    'ticket #21033 Franks, added glbOracle here to fix a problem
    'for Oracle version, it can turn on the Insync function if there are IHRDB and IHRAudit
    'registry keys setup

    rsInsync.Open "SELECT * FROM INSYNC_RPT WHERE ID=0", gdbAdoIhr001, adOpenForwardOnly
    rsInsync.Close
    
    '7.9 - Check for table and dat file.
    If File(App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "insync.dat") Then
        isInsync = True
    Else
        isInsync = False
    End If
Else
    If File(glbDBDir & "IHRINSYNC.MDB") Then
        isInsync = True
    End If
End If

Exit Function

notInsync:
    isInsync = False
End Function

Public Function isTableInDB(xTableName) As Boolean
Dim rsTBL As New ADODB.Recordset
On Error GoTo notTab

    rsTBL.Open "SELECT * FROM " & xTableName, gdbAdoIhr001, adOpenForwardOnly
    rsTBL.Close
    isTableInDB = True
Exit Function

notTab:
    isTableInDB = False
End Function

Function isWSIBModule() As Boolean
    'WSIB Form 7 - Billable Module
    'County of Wellington (2262), Chapmans (2370), City of Sarnia (2362), City of St. Thomas (2295)
    'County of Elgin (2292), County of Frontenac (2410), County of Hastings (2263), County of Lambton (2355)
    'County of Lanark (2172), County of Oxford (2259), Samuel & Son (2382),Town of St. Marys (2384)
    'City of Kawartha Lakes (2363) Ticket #22611, Walters Inc. (2443) Ticket #23227
    'Showa Canada Inc. - Ticket #24657, Windsor-Essex County Health Unit (2333) Ticket #24832
    'County of Essex (2192) - Ticket #27282, Mississaugas of Scugog Island First Nation (2485) Ticket #28648
    'City of Kenora (2487) - Ticket #29054
    If glbCompSerial = "S/N - 9999W" Or glbCompSerial = "S/N - 2262W" Or glbCompSerial = "S/N - 2370W" Or _
        glbCompSerial = "S/N - 2362W" Or glbCompSerial = "S/N - 2295W" Or glbCompSerial = "S/N - 2292W" Or _
        glbCompSerial = "S/N - 2410W" Or glbCompSerial = "S/N - 2263W" Or glbCompSerial = "S/N - 2355W" Or _
        glbCompSerial = "S/N - 2172W" Or glbCompSerial = "S/N - 2259W" Or glbCompSerial = "S/N - 2382W" Or _
        glbCompSerial = "S/N - 2384W" Or glbCompSerial = "S/N - 2363W" Or glbCompSerial = "S/N - 2443W" Or _
        glbCompSerial = "S/N - 2454W" Or glbCompSerial = "S/N - 2333W" Or glbCompSerial = "S/N - 2192W" Or _
        glbCompSerial = "S/N - 2485W" Or glbCompSerial = "S/N - 2487W" Then
        isWSIBModule = True
    Else
        isWSIBModule = False
    End If
End Function

Function isVadim() As Boolean
Dim rsVadim As New ADODB.Recordset
Dim InterDatabaseName As String
Dim InterServerName As String
Dim InterUserName As String
Dim InterUserPassword As String

On Error GoTo notVadim


rsVadim.Open "SELECT * FROM VADIM_DATABASE ", gdbAdoIhr001, adOpenForwardOnly

isVadim = True
On Error Resume Next
If rsVadim.EOF Then
    InterDatabaseName = "Vadim"
    InterServerName = SQLServerName
    InterUserName = SQLUserName
    InterUserPassword = SQLUserPassword
Else
    InterDatabaseName = rsVadim("DATABASE_NAME") & ""
    InterServerName = rsVadim("DATABASE_SERVER") & ""
    InterUserName = rsVadim("DATABASE_USERID") & ""
    InterUserPassword = rsVadim("DATABASE_PASSWORD") & ""
End If
If InterDatabaseName <> "" Then
    gdbPayroll.CursorLocation = adUseClient
    If gdbPayroll.State = adStateOpen Then gdbPayroll.Close
    gdbPayroll.CommandTimeout = 600
    gdbPayroll.Mode = adModeReadWrite
    glbPayrollDB = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & InterUserName & ";Password=" & InterUserPassword & ";Initial Catalog=" & InterDatabaseName & ";Data Source=" & InterServerName
    gdbPayroll.Open glbPayrollDB
End If
If UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then
    glbVadimEXE = "U:\HR SYSTEMS VB6\IHRVADIM\IHRVadim.EXE" '
Else
    'Ticket #15223
    If Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRVADIM.EXE") <> "" And Dir(App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRIntegration.DAT") <> "" Then
        glbVadimEXE = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRVadim.EXE"
    Else
        isVadim = False
    End If
End If
Call AddVadimTables_Map

Exit Function

notVadim:
    isVadim = False
End Function

Function isAdvancedTracker()
Dim rsADVSetup As New ADODB.Recordset
Dim xFile
Dim ETPFile
Dim ETPath
On Error GoTo notADV
isAdvancedTracker = False
xFile = App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "etpath.dat"
If Dir(xFile) <> "" Then
    isAdvancedTracker = True
    Open xFile For Input As #1
    On Error Resume Next
    Line Input #1, ETPath
    Close #1
    ETPFile = ETPath & IIf(Left(ETPath, 1) = "\", "", "\") & "etp.mdb"
    If Dir(ETPFile) <> "" Then
        Dim SQLQ
        SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='Advanced Tracker' "
        rsADVSetup.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsADVSetup.EOF Then
            rsADVSetup.AddNew
            rsADVSetup("PARA_TYPE") = "Integration"
            rsADVSetup("PARA_CATEGORY") = "Advanced Tracker"
            rsADVSetup("PARA_NAME") = "Version_Info"
            rsADVSetup("PARA_VALUE") = "MS Access"
            rsADVSetup.Update
            
            rsADVSetup.AddNew
            rsADVSetup("PARA_TYPE") = "Integration"
            rsADVSetup("PARA_CATEGORY") = "Advanced Tracker"
            rsADVSetup("PARA_NAME") = "Database_Path"
            rsADVSetup("PARA_VALUE") = ETPath
            rsADVSetup.Update
            
            rsADVSetup.AddNew
            rsADVSetup("PARA_TYPE") = "Integration"
            rsADVSetup("PARA_CATEGORY") = "Advanced Tracker"
            rsADVSetup("PARA_NAME") = "Database_Name"
            rsADVSetup("PARA_VALUE") = "etp.mdb"
            rsADVSetup.Update
        End If
        rsADVSetup.Close
    End If
End If

Exit Function

notADV:
    isAdvancedTracker = False
End Function

Public Sub UpdVacTimeRequest(zEMPNBR, Action, Optional zOldEmpnbr)
On Error GoTo EndSub
Dim gdbESS As New ADODB.Connection
Dim SQLQ
If glbSQL Or glbOracle Then
    Set gdbESS = gdbAdoIhr001
Else
    gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
End If
If Action = "D" Then
    SQLQ = "DELETE FROM HR_VACTIMEOFF_REQ WHERE VT_EMPNBR=" & zEMPNBR
Else
    SQLQ = "UPDATE HR_VACTIMEOFF_REQ SET VT_EMPNBR=" & zEMPNBR & " WHERE VT_EMPNBR=" & zOldEmpnbr
End If
gdbESS.Execute SQLQ

Exit Sub
EndSub:

End Sub

'George added on May 11,2005 for County of Essex
Public Sub UpdTimesheetEMPNum(zEMPNBR, Action, Optional zOldEmpnbr)
On Error GoTo EndSub
Dim gdbESS As New ADODB.Connection
Dim SQLQ
If glbSQL Or glbOracle Then
    Set gdbESS = gdbAdoIhr001
Else
    gdbESS.Open Replace(glbAdoIHRDB, "IHR001", "IHRESS")
End If
If Action = "D" Then
    SQLQ = "DELETE FROM HR_TIMESHEET WHERE AD_EMPNBR=" & zEMPNBR
Else
    SQLQ = "UPDATE HR_TIMESHEET SET AD_EMPNBR=" & zEMPNBR & " WHERE AD_EMPNBR=" & zOldEmpnbr
End If
gdbESS.Execute SQLQ

Exit Sub
EndSub:

End Sub

'geo added begin
Function isGreatPlains() As String
Dim xFile

On Error GoTo notGP
isGreatPlains = False
xFile = App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "greatplains.dat"
If Dir(xFile) <> "" Then
    isGreatPlains = True
End If

Exit Function

notGP:
    isGreatPlains = False
End Function

Function isIHRPWeb() 'Ticket #20209 Franks 04/21/2011
Dim xFile
Dim retval As Boolean
Dim xFlag1 As Boolean
Dim xFlag2 As Boolean

On Error GoTo notPWeb
retval = False
xFlag1 = False
xFile = App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "IHRPWeb.exe"
If Dir(xFile) <> "" Then
    xFlag1 = True
End If
xFlag2 = False
xFile = App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "IHRPWeb.dat"
If Dir(xFile) <> "" Then
    xFlag2 = True
End If

If xFlag1 And xFlag2 Then
    retval = True
    glbPayWebEXE = App.Path & IIf(Right(App.Path, 1) = "\", "", "\") & "IHRPWeb.EXE"
End If

isIHRPWeb = retval

Exit Function

notPWeb:
    glbPayWeb = False
    
End Function

Function isMediPay() As String
Dim xFile

On Error GoTo notGP
isMediPay = False
xFile = App.Path & IIf(Left(App.Path, 1) = "\", "", "\") & "MediPay.dat"
If Dir(xFile) <> "" Then
    isMediPay = True
End If

Exit Function

notGP:
    isMediPay = False
End Function
'
'geo added end
Function isEDU() As Boolean
    Dim xFile
    
    On Error GoTo Eh
    isEDU = False
    xFile = glbIHREDU
    If Dir(xFile) <> "" Then
        isEDU = True
    End If
    
    Exit Function
    
Eh:
    isEDU = False
End Function

Public Function AllCapitalString(xStr, Optional xExcept) As Boolean
    If Not IsMissing(xExcept) Then
        xStr = Replace(xStr, xExcept, "")
    End If
    If Not IsNumeric(xStr) Then
        If xStr = UCase(xStr) Then
            AllCapitalString = True
        Else
            AllCapitalString = False
        End If
    End If
End Function

' This function returns the field number from a CSV formatted string.
' ie. CSVGet("One, Two, Three, Four", 3) would return "Three".
Public Function CSVGet(FullLine As String, FieldNum) As String
    Dim I As Integer
    Dim buf As String, OneChar As String * 1
    Dim CurPos As Integer
    Dim NumCommas As Integer
    Dim InQuote As Boolean
    
    For I = 1 To Len(FullLine)
        OneChar = Mid(FullLine, I, 1)
        If OneChar = "," And InQuote = False Then
            NumCommas = NumCommas + 1
            If NumCommas = FieldNum - 1 Then
                buf = ""
            ElseIf NumCommas = FieldNum Then
                CSVGet = Trim(buf)
                Exit Function
            End If
        ElseIf OneChar = """" Then
             InQuote = Not InQuote
        Else
            buf = buf & OneChar
        End If
    Next I
    CSVGet = Trim(buf)
End Function
Public Function CharCount(OrigString As String, _
  Chars As String, Optional CaseSensitive As Boolean = False) _
  As Long

'**********************************************
'PURPOSE: Returns Number of occurrences of a character or
'or a character sequencence within a string

'PARAMETERS:
    'OrigString: String to Search in
    'Chars: Character(s) to search for
    'CaseSensitive (Optional): Do a case sensitive search
    'Defaults to false

'RETURNS:
    'Number of Occurrences of Chars in OrigString

'EXAMPLES:
'Debug.Print CharCount("FreeVBCode.com", "E") -- returns 3
'Debug.Print CharCount("FreeVBCode.com", "E", True) -- returns 0
'Debug.Print CharCount("FreeVBCode.com", "co") -- returns 2
''**********************************************

Dim lLen As Long
Dim lCharLen As Long
Dim lAns As Long
Dim sInput As String
Dim sChar As String
Dim lCtr As Long
Dim lEndOfLoop As Long
Dim bytCompareType As Byte

sInput = OrigString
If sInput = "" Then Exit Function
lLen = Len(sInput)
lCharLen = Len(Chars)
lEndOfLoop = (lLen - lCharLen) + 1
bytCompareType = IIf(CaseSensitive, vbBinaryCompare, _
   vbTextCompare)

    For lCtr = 1 To lEndOfLoop
        sChar = Mid(sInput, lCtr, lCharLen)
        If StrComp(sChar, Chars, bytCompareType) = 0 Then _
            lAns = lAns + 1
    Next

CharCount = lAns

End Function

Public Function IsReportAuth(xEmpNo)
Dim SQLQ As String
Dim rsTemp As New ADODB.Recordset
Dim retval As Boolean
    retval = False
    'for report auth 1,2,3,4
    SQLQ = "SELECT JH_EMPNBR FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
    SQLQ = SQLQ & "AND (JH_REPTAU = " & xEmpNo & " OR JH_REPTAU2 = " & xEmpNo & " OR JH_REPTAU3 = " & xEmpNo & " OR JH_REPTAU4 = " & xEmpNo & ") "
    rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsTemp.EOF Then
        retval = True
    End If
    rsTemp.Close
    IsReportAuth = retval
    'IsReportAuth = False '??? for debug, will remove this on shipment
End Function
Public Sub CreateEmpList4ReportAuth(xEmpNo)
Dim SQLQ As String
Dim rsEListWRK As New ADODB.Recordset
Dim RSEMPLIST As New ADODB.Recordset

    gdbAdoIhr001.Execute "DELETE FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"

    'for report auth 1,2,3,4
    SQLQ = "SELECT distinct JH_EMPNBR FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
    SQLQ = SQLQ & "AND (JH_REPTAU = " & xEmpNo & " OR JH_REPTAU2 = " & xEmpNo & " OR JH_REPTAU3 = " & xEmpNo & " OR JH_REPTAU4 = " & xEmpNo & ") "
    RSEMPLIST.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    SQLQ = "SELECT * FROM HR_EMPLIST_WRK WHERE TT_WRKEMP='" & glbUserID & "'"
    rsEListWRK.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Do While Not RSEMPLIST.EOF
            rsEListWRK.AddNew
            rsEListWRK("TT_COMPNO") = "001"
            rsEListWRK("TT_EMPNBR") = RSEMPLIST("JH_EMPNBR")
            rsEListWRK("TT_WRKEMP") = glbUserID
            rsEListWRK.Update
            RSEMPLIST.MoveNext
    Loop
    rsEListWRK.Close
    RSEMPLIST.Close
    
End Sub

Public Sub GWLAuditAdd(xEmpNo, xTranType, xChgType, xEffDate, xForm, xItem, xOldVal, xNewVal, Optional xLDate, Optional xTermSEQ)
Dim rsNGSAudit As New ADODB.Recordset
Dim SQLQ As String
    SQLQ = "SELECT * FROM HR_GWL_TRAN_AUDIT WHERE 1=2"
    rsNGSAudit.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsNGSAudit.AddNew
    rsNGSAudit("MT_EMPNBR") = xEmpNo
    rsNGSAudit("MT_TRAN_TYPE") = xTranType
    rsNGSAudit("MT_CHG_TYPE") = xChgType
    rsNGSAudit("MT_EFF_DATE") = xEffDate
    rsNGSAudit("MT_FORM") = xForm
    rsNGSAudit("MT_ITEM") = xItem
    rsNGSAudit("MT_NEW_VALUE") = xNewVal
    rsNGSAudit("MT_OLD_VALUE") = xOldVal
    rsNGSAudit("MT_UPLOAD") = "N"
    rsNGSAudit("MT_LUSER") = glbUserID
    If IsMissing(xLDate) Then
        rsNGSAudit("MT_LDATE") = Date
    Else
        rsNGSAudit("MT_LDATE") = xLDate
    End If
    rsNGSAudit("MT_LTIME") = Time$
    If IsMissing(xTermSEQ) Then
        rsNGSAudit("MT_TERM_SEQ") = 0
    Else
        rsNGSAudit("MT_TERM_SEQ") = xTermSEQ
    End If

    rsNGSAudit.Update
    rsNGSAudit.Close
    'to create audit report
    'create GWL Audit view using MT_EMPNBR + TermSEQ as key
    'create HREMP UNION TERM_HREMP and use ED_EMPNBR + 0 as key for Active, ED_EMPNBR + TermSEQ as key for term
    'then this report can show employee name for both active and term

End Sub

Public Function getSamAdditionalVac(xStartDate, xAsOfDate) ', xSerMonths) 'Ticket #23385 Franks 03/21/2013
'xSerMonths - the employee crosses over the 35.99 months of service during the update, it is the 2nd From Serivce
Dim xYear, xMonth, xDay
Dim xSDate, xEDate
Dim xDat2, xEndMonth, xEndYear, xLocSerMonths
Dim retval As Double
    retval = 0
    ''If Len(xSerMonths) = 0 Then GoTo end_line
    ''If Not IsNumeric(xSerMonths) Then GoTo end_line
    ''xSerMonths = Val(xSerMonths)
    ''If xSerMonths <= 0 Then GoTo end_line
    ''
    ''xYear = Year(xAsOfDate)
    ''xEndYear = CVDate("Dec 31, " & xYear)
    ''xLocSerMonths = getSamMonthDiff(CVDate(xStartDate), CVDate(xEndYear)) 'service months by the end of this year

    'If xLocSerMonths >= xSerMonths Then '
        If IsDate(xStartDate) Then
            xYear = Year(xStartDate)
            'Jan
            xSDate = CVDate("Jan 1, " & xYear)
            xEDate = CVDate("Feb 3, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 5
            'Feb
            xSDate = CVDate("Feb 4, " & xYear)
            xEDate = CVDate("Mar 8, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 4.5
            'Mar
            xSDate = CVDate("Mar 9, " & xYear)
            xEDate = CVDate("Apr 10, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 4
            'Apr
            xSDate = CVDate("Apr 11, " & xYear)
            xEDate = CVDate("May 13, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 3.5
            'May
            xSDate = CVDate("May 14, " & xYear)
            xEDate = CVDate("Jun 16, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 3
            'Jun
            xSDate = CVDate("Jun 17, " & xYear)
            xEDate = CVDate("Jul 19, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 2.5
            'Jul
            xSDate = CVDate("Jul 20, " & xYear)
            xEDate = CVDate("Aug 21, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 2
            'Aug
            xSDate = CVDate("Aug 22, " & xYear)
            xEDate = CVDate("Sep 23, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 1.5
            'Sep
            xSDate = CVDate("Sep 24, " & xYear)
            xEDate = CVDate("Oct 26, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 1
            'Oct
            xSDate = CVDate("Oct 27, " & xYear)
            xEDate = CVDate("Nov 29, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 0.5
            'Nov
            xSDate = CVDate("Nov 30, " & xYear)
            xEDate = CVDate("Dec 31, " & xYear)
            If CVDate(xStartDate) >= CVDate(xSDate) And CVDate(xStartDate) <= CVDate(xEDate) Then retval = 0
        End If
    'End If
end_line:
    getSamAdditionalVac = retval

End Function


Public Function getSamNewhireMonthVacEnt(xNewEntitle#, xStartDate, xAsOfDate) 'Ticket #23385 Franks 03/22/2013
Dim xYear1, xMonth1, xDay1
Dim xYear2, xMonth2, xDay2
Dim retval
    retval = xNewEntitle#
    xYear1 = Year(xStartDate)
    xYear2 = Year(xAsOfDate)
    If xYear1 = xYear2 Then 'must be the same year
        xMonth1 = month(xStartDate)
        xMonth2 = month(xAsOfDate)
        If xMonth1 = xMonth2 Then 'must be the same month
            xFirstMonEnt = xNewEntitle# 'Ticket #23812 Franks 05/23/2013
            xDay1 = Day(xStartDate)
            If xDay1 >= 1 And xDay1 <= 15 Then retval = xNewEntitle#  'full vac ent for this month
            If xDay1 > 15 Then retval = xNewEntitle# / 2 'half vac ent for this month
        End If
    End If
    getSamNewhireMonthVacEnt = retval
'Logic:
'Vacation will be earned on an "Annualized Monthly" basis. User-Defined Date 1 is the date field used to determine the employee's "Month of Service".
'xStartDate - employee User Defined Date(relabel to Vacation Date)
'xAsOfDate - As of Date for this month
'For new hires only
'"   Check the employee's Vacation Entitlement Code and lookup the entitlement rule.
'"   If the hire date is between the 1st and the 15th of the month, the employee will receive the full month's vacation accrual for the first month and the Entitlement amount for each of the months remaining between the months between the User Defined Date 1 and the Vacation Period's "To Date".
'"   If the hire date is after the 15th day, the employee will receive ½ month's vacation accrual for the first month and the Entitlement amount for each of the months remaining between the months between the User Defined Date 1 and the Vacation Period's "To Date".
End Function

Public Function getSamMonthDiff(xStartDate, xAsOfDate)  'Ticket #23385 Franks 03/22/2013
Dim xYear1, xMonth1, xDay1
Dim xYear2, xMonth2, xDay2
Dim retval
    'example: xStartDate = Jan 06, 2013; xAsOfDate = Jan 01, 2013,
    retval = MonthDiff(CVDate(xStartDate), CVDate(xAsOfDate))
    If retval < 0 Then
        xYear1 = Year(xStartDate)
        xYear2 = Year(xAsOfDate)
        xMonth1 = month(xStartDate)
        xMonth2 = month(xAsOfDate)
        If xYear1 = xYear2 Then 'the same year
            If xMonth1 = xMonth2 Then 'same month
                retval = 0
            End If
        End If
    End If
    getSamMonthDiff = retval
End Function

Public Sub Samuel_Vac_Ent_Cal(empNo) 'Ticket #23386 Franks 03/22/2013
Dim rsEntitle As New ADODB.Recordset
Dim SQLQ As String
Dim lastday
Dim flglastdate As Boolean
Dim blIsLast As Boolean
Dim lngRecs As Long, pct As Long, prec As Long
Dim fglbRunTimes
Dim varStartDate As Variant
Dim xdlpAsOf

    'Annualized Monthly
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & empNo & " "
    If rsEntitle.State <> 0 Then rsEntitle.Close
    rsEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEntitle.EOF Then
        If IsNull(rsEntitle("ED_USRDAT1")) Then 'Samuel Vacation Date
            GoTo exit_line
        End If
        varStartDate = rsEntitle("ED_USRDAT1")
    End If
    rsEntitle.Close
    xdlpAsOf = CVDate("Jan 1, " & Year(Date)) 'first day of year
    
    'check all records in HRVACENT table, found the matching record, populte the rule
    'if not found then skip this function
    If Not VacEntitleRuleCheck(empNo) Then
        Exit Sub
    End If

    'loop 12 months - begin
    flglastdate = False
    lastday = MonthLastDate(CVDate(xdlpAsOf))
    If CVDate(xdlpAsOf) = CVDate(lastday) Then
        flglastdate = True
    End If
    
    'Ticket #23812 Franks 05/23/2013
    SQLQ = "UPDATE HREMP SET ED_VAC = 0 WHERE ED_EMPNBR = " & empNo
    gdbAdoIhr001.Execute SQLQ
    
    For fglbRunTimes = 1 To 12
        blIsLast = False
        If fglbRunTimes = 12 Then blIsLast = True
        'note: see PensionCalculation4Samuel to populate the rule first
        Call SamuelmodUpdateSele(empNo, xdlpAsOf, fglbRunTimes, blIsLast)
        xdlpAsOf = DateAdd("m", 1, CVDate(xdlpAsOf))

        If flglastdate Then
            xdlpAsOf = CVDate(MonthLastDate(CVDate(xdlpAsOf)))
        End If
        DoEvents

    Next
    'xdlpAsOf = DateAdd("m", -12, CVDate(xdlpAsOf))
    'loop 12 months - end

exit_line:
End Sub

Private Function SamuelmodUpdateSele(empNo, xAsOf, fglbRunTimes, Optional isLast As Boolean) 'Ticket #23386 Franks 03/22/2013
'Dim empNo As Long
Dim fglbCompMonthly%
Dim SQLQ As String
Dim snapEntitle As New ADODB.Recordset
Dim dblEntitle#, dblPrevEntitle#, dblTKEEntitle#, strDivision$
Dim strJob$, dblServiceYears#
Dim spt As Variant, varStartDate As Variant, lngRecs&
Dim dblDHours#, intWhereFit&, x%, Y%, z%, dblNewEntitle#
Dim dblFTEHours#
Dim dblNewMax#, dblEntitleUpd#, DtTm As Variant
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%
Dim prec% ', xAsOf
Dim VacpcN, VacpcO, VED_DIV, VED_PT, SQLQW1
Dim if_Entitle As Boolean, if_Vacation As Boolean
Dim xComments
Dim dblEntitleDays
Dim xSALDIST As String
Dim xTotEmpHours 'Ticket #21843 Franks 04/12/2012
Dim xoptSamuelType0, xoptSamuelType1

On Error GoTo SamuelmodUpdateSele_Err
SamuelmodUpdateSele = False

If Len(xAsOf) = 0 Then
    Exit Function
End If
'If Not CR_SnapEntitle() Then Exit Function  ' create snapEntitle (form level recordset)
Screen.MousePointer = DEFAULT

''While Not snapEntitle.EOF
    'prec% = prec% + 1
    'pct% = Int(100 * (prec% / lngRecs&))
    'MDIMain.panHelp(0).FloodPercent = pct%
    if_Entitle = False
    if_Vacation = False

    If glbCompEntVac$ = "M" Or UCase(glbCompEntVac$) = "N" Then
        fglbCompMonthly% = True
    End If

    'empNo& = snapEntitle("ED_EMPNBR")
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & empNo & " "
    If snapEntitle.State <> 0 Then snapEntitle.Close
    snapEntitle.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If snapEntitle.EOF Then
        snapEntitle.Close
        Exit Function
    End If
    
    If Not IsNull(rsEntMain("VE_ROUNDENT")) Then
        If rsEntMain("VE_ROUNDENT") Then xoptSamuelType0 = True Else xoptSamuelType1 = True
    End If
    
    If IsNull(snapEntitle("ED_VAC")) Then
        dblEntitle# = 0
    Else
        dblEntitle# = snapEntitle("ED_VAC")
    End If

    If IsNull(snapEntitle("ED_PVAC")) Then
        dblPrevEntitle# = 0
    Else
        dblPrevEntitle# = snapEntitle("ED_PVAC")
    End If

    If IsNull(snapEntitle("ED_VACT")) Then
        dblTKEEntitle# = 0
    Else
        dblTKEEntitle# = snapEntitle("ED_VACT")
    End If

    spt = snapEntitle("ED_PT")

    'If IsNull(snapEntitle(fglbWDate$)) Then GoTo lblNextRec
    If IsNull(snapEntitle("ED_USRDAT1")) Then GoTo lblNextRec

    varStartDate = snapEntitle("ED_USRDAT1")

    Dim rsJOB As New ADODB.Recordset
    If rsJOB.State <> 0 Then rsJOB.Close
    rsJOB.Open "SELECT JH_DHRS,JH_FTENUM FROM qry_JobCurrent WHERE JH_EMPNBR=" & snapEntitle("ED_EMPNBR"), gdbAdoIhr001, adOpenForwardOnly
    dblDHours# = 0
    dblFTEHours# = 0
    If Not rsJOB.EOF Then
        If IsNumeric(rsJOB("JH_DHRS")) Then dblDHours# = rsJOB("JH_DHRS")
        If IsNumeric(rsJOB("JH_FTENUM")) Then dblFTEHours# = rsJOB("JH_FTENUM")
    End If
    'rsJOB.Close - move it to the botton of 2433W section
    If glbLinamar Then dblDHours# = 8

    'xAsOf = dlpAsOf.Text
'    dblServiceYears# = (DateDiff("d", varStartDate, CVDate(xAsOf)) / 365) * 12
    If glbSamuel Then 'Ticket #23385 Franks 03/22/2013
        dblServiceYears# = getSamMonthDiff(CVDate(varStartDate), CVDate(xAsOf)) ' MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    Else
        dblServiceYears# = MonthDiff(CVDate(varStartDate), CVDate(xAsOf))
    End If
    intWhereFit& = -1
    
    If dblServiceYears# < 0 Then GoTo lblNextRec 'Ticket #23812 Franks 05/23/2013
    
    For x% = 0 To 24
        If xmedGTServ(x%) > 0 Then
            If dblServiceYears# >= CDbl(xmedLTServ(x%)) And dblServiceYears# <= CDbl(xmedGTServ(x%)) Then
                intWhereFit& = x%
                If Len(xmedEntitle(x%)) > 0 Then if_Entitle = True
                If Len(xmedVacation(x%)) > 0 Then if_Vacation = True
                Exit For
            End If
        End If
    Next x%

    If intWhereFit& = -1 Then GoTo lblNextRec  ' skip record if not in any of the ranges
    rsJOB.Close

    ' Two variables glbCompEntVac$ = "M" And glbCompEntSick$ = "M"    are 'company' level
    ' which represents if Sick and Vacation entitlements
    ' are determined on monthly basis (vs yearly) - these are stored in table hrpasco
    ' and read on system startup.

    ' In this routine we work independantly of SICK/VACATIon entitlement.
    '  fglbCompMonthly% - is the independant representation
        'of glbCompEntVac$ = "M" And glbCompEntSick$ = "M"
        'Procedure SamuelmodUpdateSele is used to set
        'fglbCompMonthly based on values it finds for global variables
        ' and what the user wants to manipulate (sick/Vac)

    'optD indicates if Entitlement entered is Daily or yearly based
    ' if daily then max entitlement is based on entitlement * hours they work.

    ' we have   Entitle = existing entitmenet (stored presently
    '           NewEntitle = amount entered onto screen = medentitle(index)
    '           EntitleUpd  = value to update record with

    If if_Entitle Then
        dblNewEntitle# = xmedEntitle(intWhereFit&)
        If glbSamuel Then 'Ticket #23385 Franks 03/22/2013
            isConYear = False 'Ticket #23812 Franks 05/23/2013
            If xoptSamuelType0 Then 'optSamuelType(0).Value Then 'Service Center
                dblNewEntitle# = getSamNewhireMonthVacEnt(dblNewEntitle#, CVDate(varStartDate), CVDate(xAsOf))
                'Ticket #23385 Franks 03/25/2013 - begin
                If month(xAsOf) = 1 Then 'first month
                    xFirstMonEnt = dblNewEntitle#
                    'isConYear = False
                Else
                    If Year(CVDate(varStartDate)) = Year(CVDate(xAsOf)) Then 'Ticket #23812 Franks 05/23/2013
                        'do not setup isConYear
                    Else
                        If Not xFirstMonEnt = dblNewEntitle# Then
                        'if one month jumps to next rule during this year, then keep all the original vac ent for all 12 months
                        'then plus additional vac days
                            dblNewEntitle# = xFirstMonEnt
                            isConYear = True
                        End If
                    End If
                End If
                'Ticket #23385 Franks 03/25/2013 - end
            End If
        End If
        dblNewMax# = 0
        If xoptD(intWhereFit&) = True Then           ' Entitlements entered in days
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012
                If xmedMax(intWhereFit&) <> 0 Then dblNewMax# = xmedMax(intWhereFit&) * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If xmedMax(intWhereFit&) <> 0 Then dblNewMax# = xmedMax(intWhereFit&) * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblDHours#
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If xoptF(intWhereFit&) = True Then
            'Ticket #22766 - KidsLink - sum up the FTE for multi positions
            If glbCompSerial = "S/N - 2433W" Or glbCompSerial = "S/N - 2430W" Then 'Kerrys Place Ticket #21843 Franks 04/12/2012
                If xmedMax(intWhereFit&) <> 0 Then dblNewMax# = xmedMax(intWhereFit&) * xTotEmpHours
                dblNewEntitle# = dblNewEntitle# * xTotEmpHours
            Else
                If xmedMax(intWhereFit&) <> 0 Then dblNewMax# = xmedMax(intWhereFit&) * dblFTEHours# * dblDHours#
                dblNewEntitle# = dblNewEntitle# * dblFTEHours# * dblDHours#
            End If
            dblEntitleUpd = dblNewEntitle
        End If
        If xoptH(intWhereFit&) = True Then
            If xmedMax(intWhereFit&) <> 0 Then dblNewMax# = xmedMax(intWhereFit&)
        End If
        If fglbCompMonthly Then
            If glbCompSerial = "S/N - 2322W" Then   'Family & Children's Services of Guelph and Wellington County
                If fglbRunTimes = 1 Then
                    dblEntitleUpd = dblNewEntitle
                Else
                    dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
                    If dblDHours# <> 0 And fglbRunTimes = 12 Then
                        dblEntitleUpd# = Round25(dblEntitleUpd# / dblDHours#) * dblDHours#
                    End If

                End If
            Else
                dblEntitleUpd# = dblEntitle# + dblNewEntitle  ' accumulate monthly values
            End If
        Else
            dblEntitleUpd = dblNewEntitle ' rollover is in other utility (to accumulate)
        End If

        If dblNewMax <> 0 Then          'only do if not zero
            If glbCompSerial = "S/N - 2423W" Then
                'Accellos Canada Inc.
                xSALDIST = "N"
                If Not IsNull(snapEntitle("ED_SALDIST")) Then
                    If UCase(snapEntitle("ED_SALDIST")) = "Y" Then
                        xSALDIST = "Y"
                    End If
                End If
                If xSALDIST = "Y" Then
                    'Ticket #18644
                    'do not use the Entitlement Maximum if Salary Distribution is "Y"
                Else
                    If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblEntitle#
                    ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
                    End If
                End If
            Else
                'Ticket #21905 - kidsLINK
                If glbCompSerial = "S/N - 2430W" Then
                    If (dblEntitle# + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblEntitle#
                    ElseIf (dblEntitleUpd + dblPrevEntitle# - dblTKEEntitle#) > dblNewMax Then
                        dblEntitleUpd = dblNewMax - (dblPrevEntitle# - dblTKEEntitle#)
                    End If
                Else
                    If dblEntitleUpd + dblPrevEntitle# > dblNewMax Then
                        dblEntitleUpd = dblNewMax - dblPrevEntitle#
                    End If
                End If
            End If
        End If

        DtTm = Now
    End If

    If if_Vacation Then
        'If glbCBrant And Len(clpCode(3).Text) > 0 And snapEntitle("ED_SECTION") >= clpCode(3).Text Then
        '    VacpcN = medVacation(intWhereFit&) + dblEntitle#
        'Else
            VacpcN = xmedVacation(intWhereFit&)
        'End If
        VacpcO = snapEntitle("ED_VACPC")
        VED_DIV = snapEntitle("ED_DIV")
        VED_PT = snapEntitle("ED_PT")
        If IsNumeric(xmedVacation(intWhereFit&)) Then snapEntitle("ED_VACPC") = xmedVacation(intWhereFit&)

    End If

Contd_Mitchell:
    If if_Entitle Then

        If isLast And glbCompSerial = "S/N - 2376W" Then '#9536 on Oct 21,2005 George
            If dblDHours# <> 0 Then
                dblEntitleDays = dblEntitleUpd / dblDHours#
                dblEntitleDays = Round((dblEntitleDays / 0.25 + 0.1), 0) * 0.25 ' round to 1/4 days
                dblEntitleUpd = dblEntitleDays * dblDHours#
            Else
                dblEntitleUpd = dblEntitleUpd
            End If
        ElseIf isLast And glbSamuel Then 'Ticket #23385 Franks 03/21/2013
            If xoptSamuelType0 Then 'optSamuelType(0).Value Then 'Service Center
                If dblDHours# <> 0 Then
                    dblEntitleDays = dblEntitleUpd / dblDHours#
                Else
                    dblEntitleUpd = dblEntitleUpd
                End If
                If isConYear Then 'Ticket #23385 Franks 03/25/2013
                    dblEntitleDays = dblEntitleDays + getSamAdditionalVac(CVDate(varStartDate), CVDate(xAsOf)) ', xmedLTServ(1))
                End If
                dblEntitleDays = Round((dblEntitleDays / 0.5 + 0.01), 0) * 0.5 ' round to 1/2 days
                dblEntitleUpd = dblEntitleDays * dblDHours#
            End If
            If xoptSamuelType1 Then 'optSamuelType(1).Value Then 'Non Service Center
                If dblDHours# <> 0 Then
                    dblEntitleDays = dblEntitleUpd / dblDHours#
                    dblEntitleDays = Round((dblEntitleDays / 0.5 + 0.01), 0) * 0.5 ' round to 1/2 days
                    dblEntitleUpd = dblEntitleDays * dblDHours#
                Else
                    dblEntitleUpd = dblEntitleUpd
                End If
            End If
        ''ElseIf isLast And chkRound.Visible = True And chkRound Then
        ''    'Round the final entitlement
        ''    If dblDHours# <> 0 And optH(intWhereFit&) = False Then
        ''        dblEntitleDays = dblEntitleUpd / dblDHours#
        ''        dblEntitleDays = Round(dblEntitleDays, 0)
        ''        dblEntitleUpd = dblEntitleDays * dblDHours#
        ''    Else
        ''        dblEntitleUpd = Round(dblEntitleUpd, 0)
        ''    End If
        ''Else
        ''    If glbCompEntVac$ = "M" And chkRound.Visible = True And chkRound Then
        ''        'If month(dlpDateRange(1).Text) = month(dlpAsOf.Text) And Year(dlpDateRange(1).Text) = Year(dlpAsOf.Text) Then
        ''            'Round the final entitlement
        ''            If dblDHours# <> 0 And optH(intWhereFit&) = False Then
        ''                dblEntitleDays = dblEntitleUpd / dblDHours#
        ''                dblEntitleDays = Round(dblEntitleDays, 0)
        ''                dblEntitleUpd = dblEntitleDays * dblDHours#
        ''            Else
        ''                dblEntitleUpd = Round(dblEntitleUpd, 0)
        ''            End If
        ''        'Else
        ''        '    dblEntitleUpd = dblEntitleUpd       ' base entitlements sic/vacation
        ''        'End If
        ''    'Else
        ''    '    snapEntitle("ED_VAC") = dblEntitleUpd       ' base entitlements sic/vacation
        ''    End If
        End If

        'Ticket #22730
        'xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd
        xComments = "Current Vac. Ent. Chg from " & snapEntitle("ED_VAC") & " to " & dblEntitleUpd & ". OS: " & (IIf(IsNull(snapEntitle("ED_PVAC")), 0, snapEntitle("ED_PVAC")) + IIf(IsNull(snapEntitle("ED_VAC")), 0, snapEntitle("ED_VAC"))) - IIf(IsNull(snapEntitle("ED_VACT")), 0, snapEntitle("ED_VACT"))

        'Hemu - Ticket #11925 - Changed the Accrual Date from Effective Date to Entitlement Start Date
        'because otherwise it will not update Vadim until the date arrives in case it's not same as the
        'Entitlement Start Date.
        'Call Append_Accrual(EmpNo&, "VAC", dlpAsOf, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
        If fglbCompMonthly Then
            Call Append_Accrual(empNo, "VAC", xAsOf, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
        Else
            Call Append_Accrual(empNo, "VAC", glbEntPeriodFrom, dblEntitleUpd - Val(snapEntitle("ED_VAC") & ""), "U", xComments)
        End If

        'Hemu - 12/31/2003 Begin - Ticket #5348 - City of Chatham-Kent
        If (glbCompSerial = "S/N - 2188W" Or glbCompSerial = "S/N - 2228W") And month(CVDate(xAsOf)) = 12 Then
            snapEntitle("ED_VAC") = Round(dblEntitleUpd, 0)      ' base entitlements sic/vacation
        Else
            snapEntitle("ED_VAC") = dblEntitleUpd       ' base entitlements sic/vacation
        End If
        'Hemu - 12/31/2003 End

        'Added by bryan 13/Jun/06 Ticket#10916
        If glbCompSerial <> "S/N - 2380W" Then  'Ticket #13979 - Don't update for VitalAire - using Annual Vacation Entitlement screen to store the value to ED_ANNVAC
            snapEntitle("ED_ANNVAC") = snapEntitle("ED_VAC")
        End If
    End If
    snapEntitle.Update

    If if_Vacation Then
        Dim auVpcn As Double
        If glbCompSerial = "S/N - 2350W" Then 'Listowel Ticket#12299
            auVpcn = Val(Format(VacpcN)) * 100
        Else
            auVpcn = Val(Format(VacpcN))
        End If
        SQLQW1 = "INSERT INTO HRAUDIT (AU_TYPE,AU_NEWEMP,AU_EMPNBR,AU_VACPC,AU_OLDVAC, "
        SQLQW1 = SQLQW1 & "AU_DIVUPL,AU_PTUPL,AU_LDATE,AU_LTIME,AU_UPLOAD,AU_LUSER) "

        SQLQW1 = SQLQW1 & " VALUES('M','N'," & empNo & "," & auVpcn & "," & Val(Format(VacpcO))
        SQLQW1 = SQLQW1 & ",'" & VED_DIV & "','" & VED_PT & "', "
        SQLQW1 = SQLQW1 & Date_SQL(Date) & ", '"

        SQLQW1 = SQLQW1 & Time$ & "', "
        SQLQW1 = SQLQW1 & "'N', "
        SQLQW1 = SQLQW1 & "'" & glbUserID & "'"
        SQLQW1 = SQLQW1 & ")"
        gdbAdoIhr001X.Execute SQLQW1
    End If
    Dim xKey
    ''xKey = snapEntitle("ED_EMPNBR")
    '''xKey = xKey & "|" & Format(snapEntitle("ED_EFDATE"), "dd-mmm-yyyy")
    '''xKey = xKey & "|" & Format(snapEntitle("ED_ETDATE"), "dd-mmm-yyyy")
    ''xKey = xKey & "|" & Format(dlpDateRange(0), "dd-mmm-yyyy")
    ''xKey = xKey & "|" & Format(dlpDateRange(1), "dd-mmm-yyyy")
    ''xKey = xKey & "|VAC"
    ''xKey = xKey & "|" & dblEntitleUpd
    ''xKey = xKey & "|" & Format(xAsOf.Text, "dd-mmm-yyyy")  'Transaction Date
    ''Call Entitlements_Master_Integration(xKey, empNo) 'George added for Advance Tracker

lblNextRec:
''    snapEntitle.MoveNext
''    DoEvents
''Wend
SamuelmodUpdateSele = True
'MDIMain.panHelp(0).FloodType = 0 'Ticket #23812 Franks 05/23/2013
'gdbAdoIhr001.CommitTrans

'fTablHREMP.Close

snapEntitle.Close

Screen.MousePointer = DEFAULT

Exit Function

SamuelmodUpdateSele_Err:
'These errors are:
'13=type mismatch
'94=invalid use of null
'3018=couln't find field 'item'
If Err = 13 Or Err = 94 Or Err = 3018 Then
   ' MsgBox "Err:" & Str(Err) & Chr(10) & Error$ & Chr(10) & " SamuelmodUpdateSele" & Chr(10) & "FORM:FUENTITL.FRM"
    'commented out by RAUBREY 5/20/97
    Err = 0
    Resume Next
End If

Screen.MousePointer = DEFAULT
'glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "UpdateEntitle", "HR_EMP", "edit/Add")
Screen.MousePointer = DEFAULT
'If gintRollBack% = False Then
'    'Rollback
'    Resume Next
'Else
'    Unload Me
'End If
End Function

Public Sub PopulateVacEntRules()  'Ticket #23386 Franks 03/22/2013
Dim SQLQ, xOrder, nOrder, aa, SQLQW, glbiOneWhere
Dim rsVE As New ADODB.Recordset
Dim x
For x = 0 To 24
    xmedLTServ(x) = ""
    xmedGTServ(x) = ""
    xmedEntitle(x) = ""
    xmedMax(x) = ""
    xmedVacation(x) = ""
    xoptD(24) = ""
    xoptH(24) = ""
    xoptF(24) = ""
Next

If Not rsEntMain.EOF Then
    SQLQ = "SELECT * FROM HRVACENT WHERE (1=1) "
    If IsNull(rsEntMain("VE_DIV")) Then
        SQLQ = SQLQ & " AND VE_DIV IS NULL"
    Else
        If Len(rsEntMain("VE_DIV")) > 0 Then SQLQ = SQLQ & " AND VE_DIV = '" & rsEntMain("VE_DIV") & "'"
    End If
    If IsNull(rsEntMain("VE_DEPT")) Then
        SQLQ = SQLQ & " AND VE_DEPT IS NULL"
    Else
        If Len(rsEntMain("VE_DEPT")) > 0 Then SQLQ = SQLQ & " AND VE_DEPT = '" & rsEntMain("VE_DEPT") & "'"
    End If
    If IsNull(rsEntMain("VE_ORG")) Then
        SQLQ = SQLQ & " AND VE_ORG IS NULL"
    Else
        If Len(rsEntMain("VE_ORG")) > 0 Then SQLQ = SQLQ & " AND VE_ORG = '" & rsEntMain("VE_ORG") & "'"
    End If
    If IsNull(rsEntMain("VE_LOC")) Then
        SQLQ = SQLQ & " AND VE_LOC IS NULL"
    Else
        If Len(rsEntMain("VE_LOC")) > 0 Then SQLQ = SQLQ & " AND VE_LOC = '" & rsEntMain("VE_LOC") & "'"
    End If
    If IsNull(rsEntMain("VE_SECTION")) Then
        SQLQ = SQLQ & " AND VE_SECTION IS NULL"
    Else
        If Len(rsEntMain("VE_SECTION")) > 0 Then SQLQ = SQLQ & " AND VE_SECTION = '" & rsEntMain("VE_SECTION") & "'"
    End If
    If Not IsNull(rsEntMain("VE_EDATE")) Then
        SQLQ = SQLQ & " AND VE_EDATE = " & Date_SQL(rsEntMain("VE_EDATE"))
    End If
    If IsNull(rsEntMain("VE_EMP")) Then
        SQLQ = SQLQ & " AND VE_EMP IS NULL"
    Else
        If Len(rsEntMain("VE_EMP")) > 0 Then SQLQ = SQLQ & " AND VE_EMP = '" & rsEntMain("VE_EMP") & "'"
    End If
    If IsNull(rsEntMain("VE_PT")) Then
        SQLQ = SQLQ & " AND VE_PT IS NULL"
    Else
        If Len(rsEntMain("VE_PT")) > 0 Then SQLQ = SQLQ & " AND VE_PT = '" & rsEntMain("VE_PT") & "' "
    End If
    If IsNull(rsEntMain("VE_GRPCD")) Then
        SQLQ = SQLQ & " AND VE_GRPCD IS NULL"
    Else
        If Len(rsEntMain("VE_GRPCD")) > 0 Then SQLQ = SQLQ & " AND VE_GRPCD = '" & rsEntMain("VE_GRPCD") & "'"
    End If
    If Not IsNull(rsEntMain("VE_FRDATE")) Then
        SQLQ = SQLQ & " AND VE_FRDATE = " & Date_SQL(rsEntMain("VE_FRDATE"))
    End If
    If Not IsNull(rsEntMain("VE_TODATE")) Then
        SQLQ = SQLQ & " AND VE_TODATE = " & Date_SQL(rsEntMain("VE_TODATE"))
    End If
    
    SQLQ = SQLQ & " Order By VE_DIV,VE_DEPT,VE_ORG, VE_EDATE,VE_EMP,VE_PT,VE_LOC,VE_SECTION,VE_ORDER "
    rsVE.Open SQLQ, gdbAdoIhr001, adOpenKeyset
    
    Do While Not rsVE.EOF
        xOrder = rsVE("VE_ORDER")
        nOrder = Format(Val(xOrder), "##0") - 1
        If Not (nOrder < 0 Or nOrder > 24) Then
            If Not IsNull(rsVE("VE_BMONTH")) Then xmedLTServ(nOrder) = rsVE("VE_BMONTH")
            If Not IsNull(rsVE("VE_EMONTH")) Then xmedGTServ(nOrder) = rsVE("VE_EMONTH")
            If Not IsNull(rsVE("VE_ENTITLE")) Then xmedEntitle(nOrder) = rsVE("VE_ENTITLE")
            If Not IsNull(rsVE("VE_MAX")) Then xmedMax(nOrder) = rsVE("VE_MAX")
            If Not IsNull(rsVE("VE_PCT")) Then xmedVacation(nOrder) = rsVE("VE_PCT")
            If Not IsNull(rsVE("VE_TYPE")) Then
                If rsVE("VE_TYPE") = "D" Then xoptD(nOrder) = True
                If rsVE("VE_TYPE") = "H" Then xoptH(nOrder) = True
                If rsVE("VE_TYPE") = "F" Then xoptF(nOrder) = True
            End If
        End If
        rsVE.MoveNext
    Loop
    rsVE.Close
End If

End Sub

Public Function VacEntitleRuleCheck(xEmpNo) 'Ticket #23386 Franks 03/22/2013
'Dim rsEntMain As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim SQLQ As String
Dim recCount, I
    VacEntitleRuleCheck = False
    SQLQ = "SELECT DISTINCT VE_DIV,VE_DEPT,VE_ORG,VE_LOC,VE_SECTION,VE_EDATE,VE_EMP,VE_PT,VE_GRPCD,VE_FRDATE,VE_TODATE, VE_MANUAL "
    If glbSamuel Then 'Ticket #23385 Franks 03/21/2013
        SQLQ = SQLQ & ",VE_ROUNDENT "
    End If
    SQLQ = SQLQ & "FROM HRVACENT "

    If rsEntMain.State <> 0 Then rsEntMain.Close
    rsEntMain.Open SQLQ, gdbAdoIhr001, adOpenStatic
    I = 0
    recCount = 0
    If Not rsEntMain.EOF Then
        recCount = rsEntMain.RecordCount
    End If
    Do While Not rsEntMain.EOF
        'spShow.FloodPercent = (I / recCount) * 100
        I = I + 1
        DoEvents

        'check if this employee is in the this group
        'found it then go to the end line
        SQLQ = "SELECT * FROM HREMP "
        SQLQ = SQLQ & " WHERE (1=1) "
        If Not IsNull(rsEntMain("VE_DIV")) Then
            If Len(Trim((rsEntMain("VE_DIV")))) > 0 Then
                SQLQ = SQLQ & " AND ED_DIV = '" & Trim(rsEntMain("VE_DIV")) & "'"
            End If
        End If
        If Not IsNull(rsEntMain("VE_DEPT")) Then
            If Len(Trim((rsEntMain("VE_DEPT")))) > 0 Then
            SQLQ = SQLQ & " AND ED_DEPTNO = '" & Trim(rsEntMain("VE_DEPT")) & "'"
            End If
        End If
        If Not IsNull(rsEntMain("VE_ORG")) Then
            If Len(Trim((rsEntMain("VE_ORG")))) > 0 Then
                SQLQ = SQLQ & " AND ED_ORG = '" & Trim(rsEntMain("VE_ORG")) & "'"
            End If
        End If
        If Not IsNull(rsEntMain("VE_EMP")) Then
            If Len(Trim((rsEntMain("VE_EMP")))) > 0 Then
                SQLQ = SQLQ & " AND ED_EMP = '" & Trim(rsEntMain("VE_EMP")) & "'"
            End If
        End If
        If Not IsNull(rsEntMain("VE_PT")) Then
            If Len(Trim((rsEntMain("VE_PT")))) > 0 Then
                SQLQ = SQLQ & " AND ED_PT = '" & Trim(rsEntMain("VE_PT")) & "' "
            End If
        End If
        If Not IsNull(rsEntMain("VE_GRPCD")) Then
            If Len(Trim((rsEntMain("VE_GRPCD")))) > 0 Then
                SQLQ = SQLQ & " AND ED_EMPNBR IN "
                SQLQ = SQLQ & " (SELECT JH_EMPNBR FROM qry_JobCurrent "
                SQLQ = SQLQ & " WHERE JB_GRPCD = '" & Trim(rsEntMain("VE_GRPCD")) & "') "
            End If
        End If
        'Frank 05/13/2004 Ticket# 6176
        If Not IsNull(rsEntMain("VE_SECTION")) Then
            If Len(Trim((rsEntMain("VE_SECTION")))) > 0 Then
                SQLQ = SQLQ & " AND ED_SECTION = '" & Trim(rsEntMain("VE_SECTION")) & "'"
            End If
        End If
        If Not IsNull(rsEntMain("VE_LOC")) Then
            If glbCompSerial = "S/N - 2382W" Then  'Samuel  - Ticket #18235
                If Len(rsEntMain("VE_LOC")) > 0 Then SQLQ = SQLQ & " AND ED_VADIM1 = '" & rsEntMain("VE_LOC") & "' "
            Else
                If Len(rsEntMain("VE_LOC")) > 0 Then SQLQ = SQLQ & " AND ED_LOC = '" & rsEntMain("VE_LOC") & "' "
            End If
        End If
        
        SQLQ = SQLQ & " AND ED_EMPNBR = " & xEmpNo & " "
        If rsTemp.State <> 0 Then rsTemp.Close
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then 'the rule match this employee
            If Not IsNull(rsEntMain("VE_FRDATE")) And Not IsNull(rsEntMain("VE_TODATE")) Then
                If Len(Trim((rsEntMain("VE_FRDATE")))) > 0 And Len(Trim((rsEntMain("VE_TODATE")))) > 0 Then
                    glbEntPeriodFrom = rsEntMain("VE_FRDATE")
                    glbEntPeriodTo = rsEntMain("VE_TODATE")
               End If
            End If
            
            'populate 1 to 24 service rules
            'Call PopulatePenEntRules(rsEntMain)
            Call PopulateVacEntRules

            VacEntitleRuleCheck = True
            GoTo end_line
        End If
        rsTemp.Close

        rsEntMain.MoveNext
    Loop
end_line:

    'rsEntMain.Close
End Function
Private Function Round25(xNumb) 'Ticket #23386 Franks 03/22/2013
Dim xInteger, xDecimal, xDecTmp
    xInteger = Int(xNumb)
    xDecimal = xNumb - xInteger
    xDecTmp = 0
    If xDecimal >= 0 And xDecimal < 0.25 Then
        xDecTmp = 0
    End If
    If xDecimal >= 0.25 And xDecimal < 0.75 Then
        xDecTmp = 0.5
    End If
    If xDecimal >= 0.75 Then
        xDecTmp = 1
    End If
'    If xDecimal > 0 And xDecimal <= 0.5 Then
'        xDecTmp = 0.5
'    End If
'    If xDecimal > 0.5 Then
'        xDecTmp = 1
'    End If
    
    Round25 = xInteger + xDecTmp
End Function

Public Function Get_CourseCode_Master_Data(xCourseType, xCourseCode) As Boolean
    Dim rsCrsCodeMst As New ADODB.Recordset
    Dim SQLQ As String
          
    Get_CourseCode_Master_Data = False
    
    'If Course Code Master record is found then retrieve in global variables
    SQLQ = "SELECT * FROM HR_COURSECODE_MASTER WHERE ES_CRSCODE = '" & xCourseCode & "'"
    If glbWFC Or Len(xCourseType) = 0 Then     'Ticket #28121 - when Course Type is blank
    Else
        SQLQ = SQLQ & " AND ES_CTYPE = '" & xCourseType & "'"
    End If
    rsCrsCodeMst.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsCrsCodeMst.EOF Then
        glbCrsCode = rsCrsCodeMst("ES_CRSCODE")
        glbCrsCodeDesc = GetTABLDesc("ESCD", rsCrsCodeMst("ES_CRSCODE"))

        glbCrsCodeStrArr(1) = IIf(IsNull(rsCrsCodeMst("ES_CTYPE")), "", rsCrsCodeMst("ES_CTYPE"))  'Course Type
        glbCrsCodeStrArr(2) = IIf(IsNull(rsCrsCodeMst("ES_COORDINATED")), "", rsCrsCodeMst("ES_COORDINATED")) 'Co-Ordinated By
        glbCrsCodeStrArr(3) = IIf(IsNull(rsCrsCodeMst("ES_COMPANYNAME")), "", rsCrsCodeMst("ES_COMPANYNAME")) '
        glbCrsCodeStrArr(4) = IIf(IsNull(rsCrsCodeMst("ES_TRAINNER")), "", rsCrsCodeMst("ES_TRAINNER")) 'txtTrainerName.Text 'Trainer Name
        glbCrsCodeStrArr(5) = IIf(IsNull(rsCrsCodeMst("ES_HOURS")), "", rsCrsCodeMst("ES_HOURS")) 'txtCourseHRS.Text 'Course Hours
        glbCrsCodeStrArr(6) = IIf(IsNull(rsCrsCodeMst("ES_TBEMP")), "", rsCrsCodeMst("ES_TBEMP")) 'medEECont(0).Text 'Employee $
        glbCrsCodeStrArr(7) = IIf(IsNull(rsCrsCodeMst("ES_OTHER")), "", rsCrsCodeMst("ES_OTHER")) 'medEECont(2).Text 'Other Expenses $
        glbCrsCodeStrArr(8) = IIf(IsNull(rsCrsCodeMst("ES_TBCO")), "", rsCrsCodeMst("ES_TBCO")) 'medEECont(1).Text 'Employer $
        glbCrsCodeStrArr(9) = IIf(IsNull(rsCrsCodeMst("ES_ACCOM")), "", rsCrsCodeMst("ES_ACCOM")) 'medEECont(3).Text 'Accommodation $
        glbCrsCodeStrArr(10) = IIf(IsNull(rsCrsCodeMst("ES_LEARNING")), "", rsCrsCodeMst("ES_LEARNING")) 'medEECont(4).Text 'Learning Material $
        glbCrsCodeStrArr(11) = IIf(IsNull(rsCrsCodeMst("ES_EMPCUR")), "", rsCrsCodeMst("ES_EMPCUR")) 'clpEmpCur.Text 'Currency
        glbCrsCodeStrArr(12) = IIf(IsNull(rsCrsCodeMst("ES_OTCUR")), "", rsCrsCodeMst("ES_OTCUR")) 'clpOherCur.Text 'Currency
        glbCrsCodeStrArr(13) = IIf(IsNull(rsCrsCodeMst("ES_EMPLOYCUR")), "", rsCrsCodeMst("ES_EMPLOYCUR")) 'clpEmployerCur.Text 'Currency
        glbCrsCodeStrArr(14) = IIf(IsNull(rsCrsCodeMst("ES_ACOMCUR")), "", rsCrsCodeMst("ES_ACOMCUR")) 'clpAcomCur.Text 'Currency
        glbCrsCodeStrArr(15) = IIf(IsNull(rsCrsCodeMst("ES_LEARNINGCUR")), "", rsCrsCodeMst("ES_LEARNINGCUR")) 'clpLearnCur.Text 'Currency
        glbCrsCodeStrArr(16) = IIf(IsNull(rsCrsCodeMst("ES_TOTCUR")), "", rsCrsCodeMst("ES_TOTCUR")) 'clpTotCur.Text 'Currency
        glbCrsCodeStrArr(17) = "*" 'Flag
        'Ticket #24708 Franks 11/27/2013
        'glbCrsCodeStrArr(18) = IIf(IsNull(rsCrsCodeMst("ES_CONDUCT")), "", rsCrsCodeMst("ES_CONDUCT")) 'Conducted By
        glbCrsCodeStrArr(18) = IIf(IsNull(rsCrsCodeMst("ES_COORDINATED")), "", rsCrsCodeMst("ES_COORDINATED")) 'Coordinated By
        glbCrsCodeStrArr(19) = IIf(IsNull(rsCrsCodeMst("ES_CEUTYPE")), "", rsCrsCodeMst("ES_CEUTYPE")) 'CEU Type
        glbCrsCodeStrArr(20) = IIf(IsNull(rsCrsCodeMst("ES_METHODUSED")), "", rsCrsCodeMst("ES_METHODUSED")) 'Method Used
        
        Get_CourseCode_Master_Data = True
    Else
        'Course Code Master record not found
        Get_CourseCode_Master_Data = False
    End If
    rsCrsCodeMst.Close
    Set rsCrsCodeMst = Nothing
    
End Function

Public Sub AuditFutureDataDele(xEmpNo, xTermDate) 'Ticket #24859 Franks 01/15/2014
Dim SQLQ As String
    If IsDate(xTermDate) Then
        SQLQ = "DELETE FROM HRAUDIT WHERE AU_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND AU_LDATE > " & Date_SQL(xTermDate) & " "
        gdbAdoIhr001.Execute SQLQ
    End If
End Sub

Public Sub FamilyDayPositionSalary(xNextEmpNo, xPayID, xOldEmpNo) 'Family Day Ticket #24729 01/20/2014
Dim SQLQ As String
    Call UnloadFrms
    Unload frmSecuCopy
    Unload frmEEBASIC
    SQLQ = "UPDATE HREMP SET ED_BADGEID = '" & xPayID & "' WHERE ED_EMPNBR = " & xOldEmpNo
    gdbAdoIhr001.Execute SQLQ
    SQLQ = "UPDATE HREMP SET ED_USER_TEXT2 = 'Lock' WHERE ED_EMPNBR = " & xNextEmpNo
    gdbAdoIhr001.Execute SQLQ
    
    Call FamilyDayEphotoSync(xOldEmpNo)
    
    Call FoundLocEEFind(xNextEmpNo)
    If gSec_Inq_Position Then
        frmEPOSITION.Show
        frmEPOSITION.ZOrder 0
    Else
        MsgBox "You Do Not Have Authority For Employee Position"
    End If
End Sub

Public Function getFamilyDayUpdateRight(xUpdateRight, xEmpNo)
Dim rs As New ADODB.Recordset
Dim SQLQ As String
Dim retval As Boolean
    retval = xUpdateRight
    If glbtermopen Then
        SQLQ = "SELECT ED_EMPNBR, ED_USER_TEXT2,TERM_SEQ FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq & " "
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_USER_TEXT2 FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    End If
    SQLQ = SQLQ & "AND ED_USER_TEXT2 = 'Lock'"
    rs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rs.EOF Then
        'Lock found, then HREMP data can't be changed
        retval = False
    End If
    rs.Close
    getFamilyDayUpdateRight = retval
End Function

Public Sub FamilyDayEphotoSync(xMainEmpNo) 'Family Day Ticket #24729 01/20/2014
Dim FldName, FldNameNew
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer, iRecordCounter As Long
Dim SQLQ As String
Dim x2ndEmpID, x2ndPayID, xSIN
Dim xSkip As Boolean
    
    x2ndEmpID = FamilyDay2ndEmpNo(xMainEmpNo)
    If x2ndEmpID = 0 Then
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HR_PHOTO WHERE PT_EMPNBR = " & xMainEmpNo
    rstSource.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rstSource.EOF Then
        SQLQ = "SELECT * FROM HR_PHOTO WHERE PT_EMPNBR = " & x2ndEmpID & " "
        rstDestination.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rstDestination.EOF Then ' not found
            rstDestination.AddNew
            rstDestination("PT_EMPNBR") = x2ndEmpID
        End If
        For iFieldCounter = 0 To rstSource.Fields.count - 1
            FldName = rstSource.Fields(iFieldCounter).name
            xSkip = False
            If FldName = "PT_ID" Then xSkip = True
            If FldName = "PT_EMPNBR" Then xSkip = True
            If Not xSkip Then
                rstDestination(FldName) = rstSource(FldName)
            End If
        Next
        rstDestination.Update
        'End If
        rstDestination.Close
    End If
    rstSource.Close

End Sub

Public Sub FamilyDayEmpSync(xMainEmpNo) 'Family Day Ticket #24729 01/20/2014
Dim FldName, FldNameNew
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer, iRecordCounter As Long
Dim SQLQ As String
Dim x2ndEmpID, x2ndPayID, xSIN
Dim xSkip As Boolean
    
    x2ndEmpID = FamilyDay2ndEmpNo(xMainEmpNo)
    If x2ndEmpID = 0 Then
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xMainEmpNo
    rstSource.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rstSource.EOF Then
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & x2ndEmpID & " "
        rstDestination.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rstDestination.EOF Then 'found
            For iFieldCounter = 0 To rstSource.Fields.count - 1
                FldName = rstSource.Fields(iFieldCounter).name
                xSkip = False
                'Demo screen
                If FldName = "ED_ID" Then xSkip = True
                If FldName = "ED_EMPNBR" Then xSkip = True
                If FldName = "ED_PAYROLL_ID" Then xSkip = True
                If FldName = "ED_BADGEID" Then xSkip = True
                'Status/Date screen
                If FldName = "ED_EMP" Then xSkip = True
                If FldName = "ED_SFDATE" Then xSkip = True
                If FldName = "ED_STDATE" Then xSkip = True
                If FldName = "ED_EMPTYPE" Then xSkip = True
                If FldName = "ED_PT" Then xSkip = True
                If FldName = "ED_PTEDATE" Then xSkip = True
                If FldName = "ED_ORG" Then xSkip = True
                If FldName = "ED_BENEFIT_GROUP" Then xSkip = True
                If FldName = "ED_INTEL" Then xSkip = True
                If FldName = "ED_SALDIST" Then xSkip = True
                If FldName = "ED_EMAIL" Then xSkip = True
                If FldName = "ED_HIRECODE" Then xSkip = True
                If FldName = "ED_USER_TEXT1" Then xSkip = True
                If FldName = "ED_USER_TEXT2" Then xSkip = True
                If FldName = "ED_USER_NUM1" Then xSkip = True
                If FldName = "ED_USER_NUM2" Then xSkip = True
                'dates
                If FldName = "ED_DOH" Then xSkip = True
                If FldName = "ED_SENDTE" Then xSkip = True
                If FldName = "ED_EXPYEAR" Then xSkip = True
                If FldName = "ED_LTHIRE" Then xSkip = True
                If FldName = "ED_UNION" Then xSkip = True
                If FldName = "ED_FMLA" Then xSkip = True
                If FldName = "ED_FDAY" Then xSkip = True
                If FldName = "ED_LDAY" Then xSkip = True
                If FldName = "ED_OMERS" Then xSkip = True
                If FldName = "ED_USRDAT1" Then xSkip = True
                If FldName = "ED_DEPTEDATE" Then xSkip = True
                If FldName = "ED_DIVEDATE" Then xSkip = True
                If FldName = "ED_ELIGIBLE" Then xSkip = True
                If FldName = "ED_EARLYR" Then xSkip = True
                If FldName = "ED_NORMALR" Then xSkip = True
                If FldName = "ED_LATESTR" Then xSkip = True
                If Not xSkip Then
                    rstDestination(FldName) = rstSource(FldName)
                End If
            Next
            rstDestination.Update
        End If
        rstDestination.Close
    End If
    rstSource.Close
End Sub
Public Function FamilyDay1stEmpNo(x2ndEmpNo)
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim SQLQ As String
Dim x2ndPayID, xSIN
Dim retval
    retval = 0
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & x2ndEmpNo
    rstSource.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rstSource.EOF Then
        x2ndPayID = rstSource("ED_PAYROLL_ID")
        xSIN = rstSource("ED_SIN")

        'use SIN to find the 2nd Employee with the same SIN and Name
        SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_SIN = '" & xSIN & "' "
        SQLQ = SQLQ & "AND ED_BADGEID = '" & x2ndPayID & "' "
        rstDestination.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rstDestination.EOF Then 'found
            retval = rstDestination("ED_EMPNBR")
        End If

   End If
   FamilyDay1stEmpNo = retval
End Function
Public Function FamilyDay2ndEmpNo(xMainEmpNo)
Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim SQLQ As String
Dim x2ndPayID, xSIN
Dim retval
    retval = 0
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xMainEmpNo
    rstSource.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rstSource.EOF Then
        x2ndPayID = ""
        xSIN = ""
        If Not IsNull(rstSource("ED_BADGEID")) Then
            If Len(Trim(rstSource("ED_BADGEID"))) > 0 Then
                x2ndPayID = Trim(rstSource("ED_BADGEID"))
            End If
        End If
        If Not IsNull(rstSource("ED_SIN")) Then
            If Len(Trim(rstSource("ED_SIN"))) > 0 Then
                xSIN = Trim(rstSource("ED_SIN"))
            End If
        End If
        If Len(x2ndPayID) > 0 Then
            'use SIN to find the 2nd Employee with the same SIN and Name
            SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_SIN = '" & xSIN & "' "
            SQLQ = SQLQ & "AND ED_PAYROLL_ID = '" & x2ndPayID & "' "
            rstDestination.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rstDestination.EOF Then 'found
                retval = rstDestination("ED_EMPNBR")
            End If
        End If
   End If
   FamilyDay2ndEmpNo = retval
End Function

Public Sub FamilyDayAuditSync(xMainEmpNo, rstSource As ADODB.Recordset)  'Family Day Ticket #24729 01/20/2014
Dim FldName, FldNameNew
'Dim rstSource As New ADODB.Recordset, rstDestination As New ADODB.Recordset
Dim rstDestination As New ADODB.Recordset
Dim iCounter As Integer, iTableCounter As Integer, iFieldCounter As Integer, iRecordCounter As Long
Dim SQLQ As String
Dim x2ndEmpID, x2ndPayID, xSIN
Dim xSkip As Boolean
    
        x2ndEmpID = FamilyDay2ndEmpNo(xMainEmpNo)
        If x2ndEmpID = 0 Then
            Exit Sub
        End If
        If rstSource.EOF Then
            Exit Sub
        End If
        x2ndPayID = GetEmpData(x2ndEmpID, "ED_PAYROLL_ID", "")
        SQLQ = "SELECT * FROM HRAUDIT WHERE AU_EMPNBR = " & x2ndEmpID & " AND (1=2) "
        rstDestination.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        rstDestination.AddNew
        rstDestination("AU_COMPNO") = "001"
        rstDestination("AU_EMPNBR") = x2ndEmpID
        rstDestination("AU_PAYROLL_ID") = x2ndPayID
        For iFieldCounter = 0 To rstSource.Fields.count - 1
            FldName = rstSource.Fields(iFieldCounter).name
            xSkip = False
            If FldName = "AU_ID" Then xSkip = True
            If FldName = "AU_EMPNBR" Then xSkip = True
            If FldName = "AU_PAYROLL_ID" Then xSkip = True
            If FldName = "AU_BADGEID" Then xSkip = True
            If FldName = "AU_USER_TEXT2" Then xSkip = True
            If FldName = "AU_SALDIST" Then xSkip = True
            If Not xSkip Then
                rstDestination(FldName) = rstSource(FldName)
            End If
        Next
        rstDestination.Update
        
End Sub

Public Function Older_FollowUp_Records_Found(xReason) As Boolean
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String
    
    SQLQ = "SELECT * FROM HR_FOLLOW_UP "
    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = '" & xReason & "'" 'SREV'"   'SREV, PREV, EDUC
    SQLQ = SQLQ & " AND EF_COMPLETED <> 1"  'Not completed
    rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsFollowUp.EOF Then
        Older_FollowUp_Records_Found = True
    Else
        Older_FollowUp_Records_Found = False
    End If
End Function

Public Function getEmpSickOuting(xEmpNo) 'Ticket #27729 Franks 03/15/2016
    Dim rsEmp As New ADODB.Recordset
    Dim SQLQ As String
    Dim retval
    retval = 0
    SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsEmp.EOF Then
        retval = rsEmp("ED_PSICK") + rsEmp("ED_SICK") - rsEmp("ED_SICKT")
    End If
    getEmpSickOuting = retval
End Function

Public Function getShiftCodeforLinamar(xOrgCode)
    Dim rsTABL As New ADODB.Recordset
    Dim xNewCode
    xNewCode = xOrgCode
    'rsTABL.Open "SELECT TB_KEY FROM HRTABL WHERE TB_NAME='EDRG' AND TB_KEY='" & xOrgCode & "'", gdbAdoIhr001, adOpenForwardOnly
    'Ticket #29070 Franks 08/15/2016 - it should be SHFT not EDRG
    rsTABL.Open "SELECT TB_KEY FROM HRTABL WHERE TB_NAME='SHFT' AND TB_KEY='" & xOrgCode & "'", gdbAdoIhr001, adOpenForwardOnly
    If rsTABL.EOF Or rsTABL.BOF Then
        xNewCode = "ALL" & Mid(xOrgCode, 4)
    End If
    getShiftCodeforLinamar = xNewCode
End Function

Public Sub Recalculate_OTBANK_Ajax_AllEmployees(Optional xEmpNo) 'Town of Ajax 'Ticket #30402 Franks 08/02/2017
Dim rsEmp As New ADODB.Recordset
Dim rsAttend As New ADODB.Recordset
Dim rsAttendCT As New ADODB.Recordset
Dim SQLQ
Dim xENumber
Dim xTot As Long
Dim I As Long
    If IsMissing(xEmpNo) Then
        SQLQ = "SELECT * FROM HREMP "
    Else
        SQLQ = "SELECT * FROM HREMP WHERE ED_EMPNBR =" & xEmpNo
    End If
    rsEmp.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    
    If Not rsEmp.EOF Then
        rsEmp.MoveFirst
        I = 0
        xTot = rsEmp.RecordCount
        If xTot > 1 Then
            MDIMain.panHelp(0).FloodType = 1
            MDIMain.panHelp(1).Caption = " Please Wait" '
            MDIMain.panHelp(2).Caption = ""
        End If
        Do While Not rsEmp.EOF
            If xTot > 1 Then
                MDIMain.panHelp(0).FloodPercent = (I / xTot) * 100
                I = I + 1
                DoEvents
            End If
            xENumber = rsEmp("ED_EMPNBR")
            'Set ED_OTBANK to zero for the first time otherwise Null will be updated if some Value - Null
            SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_EMPNBR =" & xENumber
            gdbAdoIhr001.Execute SQLQ
    
            SQLQ = "SELECT SUM(AD_HRS) AS OT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'OT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttend.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
            
            SQLQ = "SELECT SUM(AD_HRS) AS CT_SUM FROM HR_ATTENDANCE WHERE LEFT(AD_REASON,2) = 'CT' AND AD_EMPNBR = " & rsEmp("ED_EMPNBR") & " GROUP BY AD_EMPNBR"
            rsAttendCT.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic

            If Not rsAttend.EOF Then
                If Not rsAttendCT.EOF Then
                    SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
                Else
                    SQLQ = "UPDATE HREMP SET ED_OTBANK = " & rsAttend("OT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
                End If
                gdbAdoIhr001.Execute SQLQ
            Else
                If Not rsAttendCT.EOF Then
                    SQLQ = "UPDATE HREMP SET ED_OTBANK = " & 0 - rsAttendCT("CT_SUM") & " WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
                Else
                    SQLQ = "UPDATE HREMP SET ED_OTBANK = 0 WHERE ED_EMPNBR = " & rsEmp("ED_EMPNBR")
                End If
                gdbAdoIhr001.Execute SQLQ
            End If
            rsAttend.Close
            rsAttendCT.Close
            
            rsEmp.MoveNext
        Loop
        If xTot > 1 Then
            MDIMain.panHelp(0).FloodPercent = 100
            MDIMain.panHelp(0).FloodType = 0
            MDIMain.panHelp(1).Caption = ""
            MDIMain.panHelp(2).Caption = ""
        End If
    End If
    rsEmp.Close
    
End Sub

