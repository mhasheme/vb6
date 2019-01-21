Attribute VB_Name = "NewModeWork"
Option Explicit
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Global glbMultiSelect


Public Enum UpdateStateEnum
    OPENING
    NoRecord
    NewRecord
End Enum

Private Type EditAbleType
    Addable As Boolean
    Updateble As Boolean
    Deleteble As Boolean
End Type


Public Enum UserUploadModeEnum
    UploadFormWithoutCheck = 0
    UploadFormWithCheck = 1
    SwitchForm = 2
End Enum

Public Enum REQTypeEnum
    OpenREQ = 0
    CloseREQ = 1
End Enum
Public Enum APPLookupTypeEnum
    BOTH = 0
    Active = 1
    HIRED = 2
End Enum

Public Enum RelateModeEnum
    RelateEMP
    RelatePos
    Reports
    MassChanges
    NothingRelate
    NoForm
    RelateSetUp
    RelateTermEmp
    RelateTransEmp
    RelateDIV
    RelateJobMaster 'Ticket #28118 Franks 02/01/2016
End Enum
Global glbUserUploadMode As UserUploadModeEnum
Global glbUpDownAction As Boolean
Global glbUPDTCNT
Global glbDocImpFile As String
Global glbDocNewRecord As Boolean
Global glbDocType As String
Global glbDocDesc As String

'Global frmCRecruitment As frmRecruitment
'Global frmCRequisition As frmRequisition
Public Sub set_PrintState(Enabled As Boolean)
MDIMain.MainToolBar.ButtonS("print").Enabled = Enabled
MDIMain.MainToolBar.ButtonS("preview").Enabled = Enabled
MDIMain.Timer1.Enabled = Enabled
End Sub

Public Sub set_Buttons(Optional UpdateMode As UpdateStateEnum)

Dim xForm As Form

Dim EditAble As EditAbleType
Dim UpdateRight As Boolean
Dim RelateMode As RelateModeEnum
Dim Printable As Boolean

Dim ButtonS As New ButtonsSetting

Dim x
On Error GoTo goButtonErr

Set xForm = MDIMain.ActiveForm

UpdateRight = get_UpdateRight(xForm)
EditAble = get_Editable(xForm)
RelateMode = get_RelateMode(xForm)
Printable = get_Printable(xForm)

ButtonS.Enabled("NewEmployee") = frmEEBASIC.UpdateRight
ButtonS.Enabled("Type:edit") = UpdateRight
ButtonS.Enabled("Type:mass") = UpdateRight
If Not UpdateRight Then
   ButtonS.Enabled("cancel") = False
End If
ButtonS.Enabled("close") = RelateMode <> NoForm

ButtonS.Enabled("print") = Printable
ButtonS.Enabled("preview") = Printable

If Not EditAble.Addable Then
    ButtonS.Enabled("NewRecord") = False
    ButtonS.Enabled("massadd") = False
End If
    
If Not EditAble.Deleteble Then
    ButtonS.Enabled("delete") = False
    ButtonS.Enabled("massdelete") = False
End If
If Not EditAble.Updateble Then
    ButtonS.Enabled("save") = False
    ButtonS.Enabled("cancel") = False
    ButtonS.Enabled("massupdate") = False
End If

If UpdateMode = NewRecord Then
    ButtonS.Enabled("delete") = False
    ButtonS.Enabled("NewRecord") = False 'Ticket #15759
End If

If UpdateMode = NoRecord Then
    ButtonS.Enabled("save") = False
    ButtonS.Enabled("cancel") = False
    ButtonS.Enabled("delete") = False
End If

If RelateMode = RelateEMP Or RelateMode = RelatePos Then
    ButtonS.Enabled("up") = True
    ButtonS.Enabled("down") = True
ElseIf RelateMode = RelateTermEmp Then
    ButtonS.Enabled("up") = False
    ButtonS.Enabled("down") = False
    ButtonS.Enabled("cancel") = False
Else
    ButtonS.Enabled("up") = False
    ButtonS.Enabled("down") = False
End If

If RelateMode = MassChanges Then
    ButtonS.Visible("Type:NewEmployee") = False
    ButtonS.Visible("Type:edit") = False
    MDIMain.MainToolBar.ButtonS(4).Visible = False
    ButtonS.Visible("Type:mass") = True
ElseIf RelateMode = Reports Then
    ButtonS.Visible("Type:NewEmployee") = False
    ButtonS.Visible("Type:edit") = False
    MDIMain.MainToolBar.ButtonS(4).Visible = False
    ButtonS.Visible("Type:mass") = False
ElseIf RelateMode = NoForm Then
    ButtonS.Visible("Type:NewEmployee") = True
    ButtonS.Visible("Type:edit") = False
    ButtonS.Visible("Type:mass") = False
Else
    ButtonS.Visible("Type:NewEmployee") = True
    ButtonS.Visible("Type:edit") = True
    ButtonS.Visible("Type:mass") = False
End If

'Email Icon
If gsEMAIL_SENDING Then
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then
            ButtonS.Visible("mail") = False
        Else
            ButtonS.Visible("mail") = True
        End If
    Else
        If glbTERM_ID = 0 Then
            ButtonS.Visible("mail") = False
        Else
            ButtonS.Visible("mail") = True
        End If
    End If
Else
    ButtonS.Visible("mail") = False
End If

'Help Icon
If (xForm Is Nothing) Then
    ButtonS.Visible("help") = False
Else
    If xForm.name = "frmEHSCorrective" Then
        ButtonS.Visible("help") = True
    Else
        ButtonS.Visible("help") = False
    End If
End If

'Ticket #22682 - Release 8.0
If Not (xForm Is Nothing) Then
    If (xForm.name = "frmEEBasic" Or xForm.name = "frmEContEmpDemo") And Not glbtermopen Then   'Ticket #29660 - Contract Employees
        ButtonS.Enabled("NewEmployee") = gSec_Add_NewHire
        
        If glbWFC Then ButtonS.Enabled("NewContractEmp") = gSec_Add_NewHire
    End If
End If

'Ticket #24184 Franks 12/04/2013
If glbWFC Then ButtonS.Visible("hrsoft") = True Else ButtonS.Visible("hrsoft") = False

'Ticket #29660 - Contract Employees
'If glbWFC And UCase(Left(App.Path, 10)) = "C:\SSWORK\" Then ButtonS.Visible("NewContractEmp") = True Else ButtonS.Visible("NewContractEmp") = False
If glbWFC Then ButtonS.Visible("NewContractEmp") = True Else ButtonS.Visible("NewContractEmp") = False
  

Exit Sub

goButtonErr:
MsgBox Err.Description
Resume Next
End Sub

Public Function isUpdated(ByRef frmName As Form) ', Optional ShowButton) As Boolean
Dim Msg$, VReturn%
Dim RMode As RelateModeEnum
Dim xChanged As Boolean

isUpdated = True
RMode = get_RelateMode(frmName)
If DataModifyForm(RMode) Then
    If frmName.name = "frmEBENEFITS" Then
        If InStr(frmName.cmdBens.Caption, "Beneficiary") <> 0 Then
            xChanged = isChanged(frmName)
        Else
            xChanged = frmName.isChangedBens()
        End If
    ElseIf frmName.name = "frmSLabel" Then
        xChanged = frmName.isChangedLabel()
    
    ElseIf frmName.name = "frmSEmpFlags" Then
        xChanged = frmName.isChangedLabel()
    ElseIf frmName.name = "frmEHSF9" Then
        xChanged = isChanged(frmName)
        If Not xChanged Then
            xChanged = frmName.isChangedHS()
        End If
    Else
        xChanged = isChanged(frmName)
    End If
    If xChanged Then
        Msg$ = "Do you want to save changes?"
        frmName.ZOrder 0
        VReturn% = MsgBox(Msg$, MB_YESNO, frmName.Caption)
        If VReturn% = IDYES Then  'YES save changes
            If frmName.cmdOK_Click Then
                Pause (0.5)
            Else
                isUpdated = False
            End If
        ElseIf VReturn% = IDNO Then
            Call frmName.cmdCancel_Click
        End If
    End If
End If


End Function

Public Function isChanged(ByRef frmName As Form) As Boolean
Dim varData As Variant
Dim varCtrl As Variant
Dim Ctrl As Control
Dim RMode As RelateModeEnum
isChanged = False
'On Error Resume Next
RMode = get_RelateMode(frmName)
If Not DataModifyForm(RMode) Then Exit Function

If Not (frmName.Addable Or frmName.Updateble Or frmName.Deleteble) Then Exit Function
If frmName.Data1.Recordset Is Nothing Then Exit Function
If frmName.ChangeAction = NewRecord Then isChanged = True: Exit Function
'frmName.Data1.Refresh 'added by Bryan 05-08-05 Ticket #9063
If frmName.Data1.Recordset.EOF Then Exit Function

'frmName.Data1.Refresh
For Each Ctrl In frmName
    If TypeOf Ctrl Is Label _
    Or TypeOf Ctrl Is ComboBox _
    Or TypeOf Ctrl Is CodeLookup _
    Or TypeOf Ctrl Is TextBox _
    Or TypeOf Ctrl Is DateLookup _
    Or TypeOf Ctrl Is EmployeeLookup _
    Or TypeOf Ctrl Is MaskEdBox _
    Or TypeOf Ctrl Is CheckBox _
    Or TypeOf Ctrl Is SSCheck Then
        If Not Ctrl.name = "Updstats" Then
            If Len(Ctrl.DataField) > 0 Then
                
                If frmName.name = "frmEESTATS" And Left(Ctrl.DataField, 3) = "ER_" Then
                    'Ticket #15576
                    'Check Pension Date 1 to 6 and Other Date 1 to 10
                    'These date fields exist in HREMP_OTHER
                    If Not frmName.DataOther.Recordset.EOF Then 'Added by Frank on 02/17/2010
                        varData = frmName.DataOther.Recordset(Ctrl.DataField)
                    End If
                ElseIf frmName.name = "frmEHSINJURYWF7" And Left(Ctrl.DataField, 3) = "F7_" Then
                    'There is more than one recordset in this form
                    If Not frmName.data2.Recordset.EOF Then 'Added by Hemu - 11/09/2011
                        varData = frmName.data2.Recordset(Ctrl.DataField)
                    End If
                ElseIf frmName.name = "frmEHSF9" And Left(Ctrl.DataField, 3) = "F9_" Then
                    'There is more than one recordset in this form
                    If Not frmName.data2.Recordset.EOF Then 'Added by Hemu - Ticket #21463
                        varData = frmName.data2.Recordset(Ctrl.DataField)
                    End If
                Else
                    varData = frmName.Data1.Recordset(Ctrl.DataField)
                End If
                
                Select Case VarType(varData)
                    Case vbNull Or vbEmpty
                        If Ctrl <> "" And Ctrl <> 0 Then
                            GoTo MarkChanged
                        End If
                    Case vbDate
                        If IsDate(Ctrl) Then
                            If DaysBetween(Ctrl, varData) <> 0 Then
                                GoTo MarkChanged
                            End If
                        Else
                            If Not IsNull(varData) Then GoTo MarkChanged
                        End If
                    Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency
                        If glbFrench Then
                            If Round(Ctrl, 2) <> Round(varData, 2) Then
                               GoTo MarkChanged
                            End If
                        Else
                            If Round(Val(Ctrl), 2) <> Round(varData, 2) Then
                               GoTo MarkChanged
                            End If
                        End If
                    Case vbDecimal
                        If glbOracle Then
                            If TypeOf Ctrl Is CheckBox Or TypeOf Ctrl Is SSCheck Then
                                varData = CInt(varData): varData = IIf(varData = -1, 1, varData)
                                varCtrl = CInt(Ctrl): varCtrl = IIf(varCtrl = -1, 1, varCtrl)
                                If varCtrl <> varData Then
                                    GoTo MarkChanged
                                End If
                            Else
                                If Round(Val(Ctrl), 2) <> Round(varData, 2) Then
                                   GoTo MarkChanged
                                End If
                            End If
                        Else
                            If glbFrench Then
                                If Round(Ctrl, 2) <> Round(varData, 2) Then
                                   GoTo MarkChanged
                                End If
                            Else
                                If Round(Val(Ctrl), 2) <> Round(Val(varData), 2) Then   'Add Val() for vardata because 1259.505 = 1259.5 instead of 1259.51
                                   GoTo MarkChanged
                                End If
                            End If
                        End If
                    Case vbBoolean
                        varData = CInt(varData): varData = IIf(varData = -1, 1, varData)
                        varCtrl = CInt(Ctrl): varCtrl = IIf(varCtrl = -1, 1, varCtrl)
                        If varCtrl <> varData Then
                            GoTo MarkChanged
                        End If
                    Case vbString
                        If Ctrl.name = "medPCode" Then
                            varData = Replace(varData, " ", "")
                            varData = Replace(varData, "-", "")
                            varCtrl = Replace(Ctrl, " ", "")
                            'varCtrl = Replace(Ctrl, "-", "") 'Ticket #21543 Franks 02/08/2012
                            varCtrl = Replace(varCtrl, "-", "")  'Hemu - It should be varCtrl instead of Ctrl as the value in the previous line is put in varCtrl
                        ElseIf Ctrl.name = "medPhone" Or Ctrl.name = "medTelephone" Or _
                            Ctrl.name = "medTele2" Or Ctrl.name = "medCTele" Or Ctrl.name = "medDTele" Then
                            varData = Replace(varData, "(", "")
                            varData = Replace(varData, ")", "")
                            varData = Replace(varData, "-", "")
                            varData = Replace(varData, " ", "")
                            varCtrl = Replace(Ctrl, "(", "")
                            varCtrl = Replace(varCtrl, ")", "")
                            varCtrl = Replace(varCtrl, "-", "")
                            varCtrl = Replace(varCtrl, " ", "")
                        Else
                            varCtrl = Ctrl
                        End If
                        
                        If Trim(varCtrl) <> Trim(varData) Then
                            GoTo MarkChanged
                        End If
                End Select
            End If
        End If
    End If
Next
Exit Function

MarkChanged:
    isChanged = True
End Function

Public Function get_UpdateRight(ByRef frmName As Form) As Boolean
On Error Resume Next
get_UpdateRight = False
If frmName Is Nothing Then Exit Function

get_UpdateRight = frmName.UpdateRight
End Function
Private Function get_Editable(ByRef frmName As Form) As EditAbleType
On Error Resume Next
get_Editable.Addable = False
get_Editable.Updateble = False
get_Editable.Deleteble = False

If frmName Is Nothing Then Exit Function

get_Editable.Addable = frmName.Addable
get_Editable.Updateble = frmName.Updateble
get_Editable.Deleteble = frmName.Deleteble

End Function
Public Function DataModifyForm(ByRef RMode As RelateModeEnum) As Boolean
If RMode = RelateEMP Or RMode = RelatePos Or RMode = RelateSetUp Or RMode = RelateTermEmp Or RMode = RelateJobMaster Then
    DataModifyForm = True
Else
    DataModifyForm = False
End If

End Function
Private Function get_Printable(ByRef frmName As Form) As Boolean
On Error Resume Next
get_Printable = False
If frmName Is Nothing Then Exit Function

get_Printable = frmName.Printable
End Function
Public Function get_RelateMode(ByRef frmName As Form) As RelateModeEnum
On Error Resume Next

get_RelateMode = NoForm
If frmName Is Nothing Then Exit Function

get_RelateMode = frmName.RelateMode
End Function
Public Function SetLIST(RMode As RelateModeEnum) As adodb.Recordset
Dim SQLQ
Set SetLIST = New adodb.Recordset


Select Case RMode

Case RelateEMP
    SQLQ = "Select ED_EMPNBR,ED_SURNAME, ED_FNAME,ED_COUNTRY,ED_ORG,"
    If glbLinamar Then
        SQLQ = SQLQ & "ED_REGION AS PROD_LINE,"     'Ticket #14775
        SQLQ = SQLQ & "right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3) AS EMPNBR "
    ElseIf glbOracle Then
        SQLQ = SQLQ & "ED_EMPNBR AS EMPNBR "
    Else
        SQLQ = SQLQ & "LTRIM(STR(ED_EMPNBR)) AS EMPNBR "
    End If
    
    If glbtermopen Then
        '7.9 - Scrolling down the using the Down Arrow key sometimes retrieves records with no corresponding
        'Termination record in Term_HRTRMEMP table
        'SQLQ = SQLQ & ",TERM_SEQ FROM Term_HREMP
        SQLQ = SQLQ & ",Term_HREMP.TERM_SEQ FROM Term_HREMP INNER JOIN "
        SQLQ = SQLQ & "Term_HRTRMEMP ON Term_HRTRMEMP.TERM_SEQ = Term_HREMP.TERM_SEQ "
        If glbLinamar Then
            SQLQ = SQLQ & "LEFT JOIN HR_DIVISION  ON Term_HREMP.ED_DIV=HR_DIVISION.DIV "
        End If
    Else
        SQLQ = SQLQ & " FROM HREMP"
    End If
    
    SQLQ = SQLQ & " Where " & glbSeleDeptUn

    If glbSort <> "NUMBER" Then
       If glbOracle Then
           SQLQ = SQLQ & " ORDER BY UPPER(ED_SURNAME), UPPER(ED_FNAME)"
       Else
           SQLQ = SQLQ & " ORDER BY ED_SURNAME, ED_FNAME"
       End If
    Else
        If glbLinamar Then
            SQLQ = SQLQ & " ORDER BY EMPNBR"
        Else
            SQLQ = SQLQ & " ORDER BY ED_EMPNBR"
        End If
    End If

    If glbtermopen Then
        SetLIST.Open SQLQ, gdbAdoIhr001X, adOpenKeyset     'term
    Else
        SetLIST.Open SQLQ, gdbAdoIhr001, adOpenKeyset      'active
    End If
Case RelatePos
    SQLQ = "SELECT JB_CODE,JB_DESCR FROM HRJOB "
    SQLQ = SQLQ & " ORDER BY JB_DESCR"
    SetLIST.Open SQLQ, gdbAdoIhr001, adOpenKeyset
Case RelateDIV
    SQLQ = "SELECT * FROM HR_DIVISION "
    SQLQ = SQLQ & " WHERE " & glbSeleDiv
    SQLQ = SQLQ & " ORDER BY Division_Name "
    SetLIST.Open SQLQ, gdbAdoIhr001, adOpenKeyset
End Select
glbUpDownAction = True
End Function


Private Sub setGLBS(RMode As RelateModeEnum, ALIST As adodb.Recordset)
Select Case RMode
    Case RelateEMP
    
        If glbtermopen Then
               
            glbTERM_ID = ALIST("ED_EMPNBR")
            glbLEE_FName = ALIST("ED_FNAME") & ""
            glbLEE_SName = ALIST("ED_SURNAME") & ""
            If glbLinamar Then
                glbLEE_ProdLine = Mid(ALIST("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", ALIST("PROD_LINE")) 'Ticket #14775
            End If
            If Not IsNull(ALIST("TERM_SEQ")) Then
                glbTERM_Seq = ALIST("TERM_SEQ")
            Else
                glbTERM_Seq = 0
            End If
            glbEmpCountry = UCase(ALIST("ED_COUNTRY"))
            glbUNION = ALIST("ED_ORG") & ""
        Else
            glbLEE_ID = ALIST("ED_EMPNBR")
            glbLEE_SName = ALIST("ED_SURNAME") & ""
            glbLEE_FName = ALIST("ED_FNAME") & ""
            If glbLinamar Then
                glbLEE_ProdLine = Mid(ALIST("PROD_LINE"), 4) & " - " & GetTABLDesc("EDRG", ALIST("PROD_LINE")) 'Ticket #14775
            End If
            glbEmpCountry = UCase(ALIST("ED_COUNTRY"))
            glbUNION = IIf(Not IsNull(ALIST("ED_ORG")), ALIST("ED_ORG"), "")
        End If
        If glbWFC And glbOnTop = "FRMESALARY" Then
            If glbtermopen Then
                If WFC_Security(glbTERM_Seq, "TERM") Then
                    Unload MDIMain.ActiveForm
                End If
            Else
                If WFC_Security(glbLEE_ID, "ACTV") Then
                    Unload MDIMain.ActiveForm
                End If
            End If
        End If
    Case RelatePos
        glbPos = ALIST("JB_CODE")
        glbPosDesc = ALIST("JB_DESCR")
    Case RelateDIV
        glbDiv = ALIST("DIV")
        glbDivDesc = ALIST("Division_Name")
End Select
End Sub
Function WFC_Security(EmpID, xType)
Dim rsTemp As New adodb.Recordset
Dim SQLQ
    WFC_Security = True
    If xType = "TERM" Then
        glbBand = ""
        SQLQ = "SELECT SH_EMPNBR,SH_BAND FROM Term_SALARY_HISTORY WHERE SH_CURRENT <>0 AND TERM_SEQ = " & EmpID
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp("SH_BAND")) Then
                glbBand = rsTemp("SH_BAND")
            End If
        End If
        rsTemp.Close
    Else
        glbBand = ""
        SQLQ = "SELECT SH_EMPNBR,SH_BAND FROM HR_SALARY_HISTORY WHERE SH_CURRENT <>0 AND SH_EMPNBR = " & EmpID
        rsTemp.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTemp.EOF Then
            If Not IsNull(rsTemp("SH_BAND")) Then
                glbBand = rsTemp("SH_BAND")
            End If
        End If
        rsTemp.Close
    End If
    
    If glbNoNONE Then
        If glbUNION = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
    End If
    If glbNoEXEC Then
        If glbUNION = "EXEC" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Screen.MousePointer = DEFAULT
            Exit Function
        End If
    End If
    If gSec_WFC_Band_Security Then
        If Len(glbBand) > 0 Then
            If InStr(1, ",A,B,C,D,E,", "," & glbBand & ",") = 0 Then
                MsgBox "You Do Not Have Authority For This Transaction"
                glbOnTop = Empty
                Screen.MousePointer = DEFAULT
                Exit Function
            End If
        End If
    End If
    WFC_Security = False
End Function
Public Sub RowUp()
Dim ALIST As New adodb.Recordset
Dim RMode As RelateModeEnum

On Error GoTo goNegotiateErr
If MDIMain.ActiveForm Is Nothing Then
    RMode = RelateEMP
Else
    RMode = get_RelateMode(MDIMain.ActiveForm)
End If
If glbLinHS Then
    RMode = RelateDIV
End If
Set ALIST = SetLIST(RMode)
If ALIST.RecordCount = 0 Then GoTo FirstRow
Select Case RMode
    Case RelateEMP
        If glbtermopen Then
              If glbTERM_Seq = 0 Then GoTo FirstRow
              ALIST.Find "ED_EMPNBR=" & glbTERM_ID
        Else
                If glbLEE_ID = 0 Then GoTo FirstRow
                ALIST.Find "ED_EMPNBR=" & glbLEE_ID
        End If
    Case RelatePos
        If glbPos = "" Then GoTo FirstRow
        ALIST.Find "JB_CODE='" & glbPos & "'"
    Case RelateDIV
        If glbDiv = "" Then GoTo FirstRow
        ALIST.Find "DIV='" & glbDiv & "'"
End Select

If Not ALIST.BOF Then ALIST.MovePrevious
If ALIST.BOF Then ALIST.MoveFirst: GoTo FirstRow
Call setGLBS(RMode, ALIST)
If RMode = RelateDIV Then
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
    RMode = RelateEMP
End If
Call ReDisplayForms(RMode)


Exit Sub
FirstRow:
    MsgBox "You have reached first row."
Exit Sub

goNegotiateErr:
MsgBox Err.Description
  
End Sub

Public Sub RowDown()
Dim ALIST As New adodb.Recordset
Dim RMode As RelateModeEnum
On Error GoTo goNegotiateErr
If MDIMain.ActiveForm Is Nothing Then
    RMode = RelateEMP
Else
    RMode = get_RelateMode(MDIMain.ActiveForm)
End If
If glbLinHS Then
    RMode = RelateDIV
End If
Set ALIST = SetLIST(RMode)
If ALIST.RecordCount = 0 Then GoTo LastRow
Select Case RMode
    Case RelateEMP
        If glbtermopen Then
            If glbTERM_ID <> 0 Then ALIST.Find "ED_EMPNBR=" & glbTERM_ID
        Else
            If glbLEE_ID <> 0 Then ALIST.Find "ED_EMPNBR=" & glbLEE_ID
        End If
    Case RelatePos
        If glbPos <> "" Then ALIST.Find "JB_CODE='" & glbPos & "'"
    Case RelateDIV
        If glbDiv <> "" Then ALIST.Find "DIV='" & glbDiv & "'"
End Select

If Not ALIST.EOF Then ALIST.MoveNext
If ALIST.EOF Then ALIST.MoveLast: GoTo LastRow

Call setGLBS(RMode, ALIST)
If RMode = RelateDIV Then
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
    RMode = RelateEMP
End If
Call ReDisplayForms(RMode)

Exit Sub
LastRow:
    MsgBox "You have reached last row."
Exit Sub

goNegotiateErr:
MsgBox Err.Description
    
End Sub

Public Sub ReDisplayForms(RMode As RelateModeEnum)
Dim xForm As Form
Dim FormRMode As RelateModeEnum
If DataModifyForm(RMode) Then
    For Each xForm In Forms
        'Debug.Print xForm.name
        Err.Number = 0
        FormRMode = get_RelateMode(xForm)
        If FormRMode = RMode Then
            On Error Resume Next
            If glbWFC Then 'Ticket #23456 Franks 04/03/2013
                If xForm.name = "frmDEPNDTS" Then
                    Unload xForm
                End If
            End If
            Call xForm.Form_Load
            If Err.Number <> 0 Then
                Call xForm.EERetrieve
                Call xForm.Display_Value
            End If
        Else
            On Error Resume Next
            If xForm.name <> "MDIMain" Then
                If xForm.MDIChild Then Unload xForm
            End If
        End If
   Next
   DoEvents
   Call MDIMain.ActiveForm.Display_Value
End If
Exit Sub
End Sub

Public Function getPosDesc(ByRef JBCODE As String) As String
Dim rsPOS As New adodb.Recordset
rsPOS.Open "SELECT JB_DESCR FROM HRJOB WHERE JB_CODE='" & JBCODE & "'", gdbAdoIhr001, adOpenForwardOnly
If rsPOS.EOF Then
Exit Function
    getPosDesc = ""
Else
    getPosDesc = rsPOS("JB_DESCR")
End If
End Function

Public Function getLanguage(EmpNumber) As String
    Dim RsLang As New adodb.Recordset
    Dim iLoop
    getLanguage = "NoLang1|NoLang2"
    iLoop = 0
    RsLang.Open "SELECT EL_LANG_SPOKEN FROM HR_LANGUAGE WHERE EL_EMPNBR=" & EmpNumber & " order by EL_LANGNO ASC", gdbAdoIhr001, adOpenForwardOnly
    If RsLang.EOF Then Exit Function
    Do While Not RsLang.EOF
        iLoop = iLoop + 1
        If Not IsNull(RsLang("EL_LANG_SPOKEN")) Then
            If iLoop = 1 Then
                getLanguage = Replace(getLanguage, "NoLang1", RsLang("EL_LANG_SPOKEN"))
            Else
                getLanguage = Replace(getLanguage, "NoLang2", RsLang("EL_LANG_SPOKEN"))
            End If
        End If
        If iLoop = 2 Then Exit Do
        RsLang.MoveNext
    Loop
    RsLang.Close
    Set RsLang = Nothing
End Function

Public Function DocumentType_Access(xDocType) As Boolean

    If Len(xDocType) > 0 Then
        Dim rs As New adodb.Recordset
        Dim strSQL As String
        Dim xWrongPos, xPos, I
        Dim xList, xShowCell, xCell
        Dim xTemplate As String
        
        DocumentType_Access = True
        
        '????Ticket #24808 -  Get User's Template if there is one to retrieve template's security profile
        xTemplate = ""
        xTemplate = Get_Template(glbUserID)
        
        
        xList = xDocType
        xWrongPos = 0
        xPos = 0
        Do While Len(xList) <> 0
            xWrongPos = xWrongPos + xPos
            xPos = InStr(xList, ",")
            If xPos = 0 Then
                xShowCell = xList
                xList = ""
            Else
                xShowCell = Left(xList, xPos - 1)
                xList = Mid(xList, xPos + 1)
            End If
            xCell = xShowCell
            
            If xTemplate = "" Or xTemplate = "TEMPLATE" Then
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(glbUserID, "'", "''") & "'"
            Else
                '????Ticket #24808 -  Retrieve template's security profile
                strSQL = "SELECT ACCESSABLE FROM HR_SECURE_DOCUMENT_TYPE WHERE USERID='" & Replace(xTemplate, "'", "''") & "'"
            End If
            strSQL = strSQL & " AND CODENAME = '" & xCell & "' AND TB_NAME='DOCT'"
            rs.Open strSQL, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
            If rs.EOF = False And rs.BOF = False Then
                If rs("ACCESSABLE") = 0 Then
                    DocumentType_Access = False
                    MsgBox "You do not have Authorization to View '" & xCell & "' Document Type attachments.", vbInformation + vbOKOnly, "Authorization Failure"
                    'SendKeys "{HOME}"
                    'For I = 1 To xWrongPos
                    '    SendKeys "{Right}"
                    'Next
                    Exit Function
                End If
            Else
                DocumentType_Access = False
                MsgBox "You do not have Authorization to View '" & xCell & "' Document Type attachments.", vbInformation + vbOKOnly, "Authorization Failure"
                'SendKeys "{HOME}"
                'For I = 1 To xWrongPos
                '    SendKeys "{Right}"
                'Next
                Exit Function
            End If
            rs.Close
            Set rs = Nothing
        Loop
    End If

End Function

Private Sub FileAttribute(xFileName, xAttribute, xTempDir)
Dim TempFile2
Dim errPoint As String

    On Error GoTo ErrHandler_FileAttrib
    
    errPoint = "Error Point FileAttrib 1"
    
    TempFile2 = Replace(Replace(xTempDir, Chr(0), "") & "\IhrDoc.Bat", "\\", "\")
    
    errPoint = "Error Point FileAttrib 2"
    
    Open TempFile2 For Output As #5
    
    errPoint = "Error Point FileAttrib 3"
    
    Print #5, "attrib " & xAttribute & " " & GetShortName(xFileName)
    
    errPoint = "Error Point FileAttrib 4"
    
    Close #5
    
    errPoint = "Error Point FileAttrib 5"
    
    Shell "cmd /c " & GetShortName(TempFile2)
    
    errPoint = "Error Point FileAttrib 6"

Exit Sub
    
ErrHandler_FileAttrib:
    MsgBox Err.Description & " - Attribute: " & xAttribute & " - " & xFileName & ": " & TempFile2 & " - " & errPoint, , "Error"
End Sub

Public Function GetShortName(ByVal sLongFileName As String) As String
    Dim lRetVal As Long, sShortPathName As String, iLen As Integer
    'Set up buffer area for API function call return
    sShortPathName = Space(255)
    iLen = Len(sShortPathName)

    'Call the function
    lRetVal = GetShortPathName(sLongFileName, sShortPathName, iLen)
    'Strip away unwanted characters.
    GetShortName = Left(sShortPathName, lRetVal)
End Function

Public Function FillMemoFile(zSQLQ, zName) ' As Long)
    
    On Error GoTo ErrHandler
    
    Dim rsPHOTO As New adodb.Recordset
    Dim byteChunk() As Byte

    Dim Offset As Long
    Dim Totalsize As Long
    Dim Remainder As Long

    Dim FieldSize As Long
    Dim FileNumber As Integer
    Const HeaderSize As Long = 78
    Const ChunkSize As Long = 100
    Dim TempFile As String
    Dim TempDir As String * 255
    Dim FileExt As String
    Dim SQLQ
    Dim errPoint As String
    
    
    GetTempPath 255, TempDir

    errPoint = "Err Point - 1"

    SQLQ = zSQLQ
    rsPHOTO.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic, adLockOptimistic
    
    If rsPHOTO.EOF Then Exit Function
    
    'Release 8.1
    'Check if the User has Security rights to view the attachment based on the Document Type Code of the attachment.
    If Not DocumentType_Access(rsPHOTO("DOCTYPE")) Then Exit Function
    
    If IsNull(rsPHOTO("FILEEXT")) Then
        FileExt = ""
        TempFile = Replace(Replace(TempDir, Chr(0), "") & "\Ihr" & zName & ".tmp", "\\", "\")
        
        errPoint = "Err Point - 2"
    Else
        FileExt = rsPHOTO("FILEEXT")
        TempFile = Replace(Replace(TempDir, Chr(0), "") & "\Ihr" & zName & "." & FileExt & "", "\\", "\")
        
        errPoint = "Err Point - 3"
    End If

    FileNumber = FreeFile
    
    errPoint = "Err Point - 4"
    
    If (Dir(TempFile)) <> "" Then
        Call FileAttribute(TempFile, "-r", TempDir)
        
        errPoint = "Err Point - 4a"
        
        Call Pause(5)
        
        Kill TempFile
        
        errPoint = "Err Point - 5"
    End If
    
    Open TempFile For Binary Access Write As FileNumber
    
    errPoint = "Err Point - 6"
    
    ReDim byteChunk(rsPHOTO("DOC").ActualSize)
    errPoint = "Err Point - 7"
    
    'Ticket #18176 - Since the .docx and xlsx are actually a zip file, when doing the following adds an
    'extra byte causing the file not to be opened in Word or Excel. By triming that extra byte seems to
    'resolve the issue. This trimmed byte does not cause any issue for normal doc, xls, or pdf file so
    'no condition has been added that it should trim for .docx and .xlsx only for now.
    'byteChunk() = rsPHOTO("DOC").GetChunk(rsPHOTO("DOC").ActualSize)
    byteChunk() = rsPHOTO("DOC").GetChunk(rsPHOTO("DOC").ActualSize - 1)
    
    errPoint = "Err Point - 8"
    
    Put FileNumber, , byteChunk()

    errPoint = "Err Point - 9"
    
    Close FileNumber
    errPoint = "Err Point - 9a"
    'Kill (TempFile)
    rsPHOTO.Close
    
    errPoint = "Err Point - 9b"
    
    'Read only
    Call FileAttribute(TempFile, "+r", TempDir)
'    TempFile2 = Replace(Replace(TempDir, Chr(0), "") & "\IhrDoc.Bat", "\\", "\")
'    FileNumber = FreeFile
'    Open TempFile2 For Output As #5
'    Print #5, "attrib +r " & GetShortName(TempFile)
'    Close #5
'    Shell "cmd /c " & GetShortName(TempFile2)
    
    errPoint = "Err Point - 10"
    
    'Open the attachment
    Shell "cmd /c " & GetShortName(TempFile)

    errPoint = "Err Point - 11"
    
    Exit Function
    
ErrHandler:
    MsgBox Err.Description & " - " & FileNumber & ": " & TempFile & " - " & errPoint, , "Error"
    
End Function

Public Function getSQL(zFormName)
    Dim xFList As String
    
    getSQL = ""
    
    If Len(glbDocKey) = 0 Then glbDocKey = 0
    
    Select Case zFormName
    Case "frmEmployeeFlags"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EMP_FLAGS", "EF_DOC,EF_FILEEXT,EF_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", EF_DOC AS DOC,EF_FILEEXT as FILEEXT, EF_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EMP_FLAGS", "EF_DOC, EF_FILEEXT,EF_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", EF_DOC AS DOC,EF_FILEEXT as FILEEXT, EF_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE='" & UCase(glbDocName) & "' AND EF_EMPNBR=" & glbLEE_ID
        End If
        'getSQL = getSQL & " AND EF_DOCKEY= " & glbDocKey & " "
        getSQL = getSQL & " AND EF_FLAGDTE = " & Date_SQL(glbEmpFlagDate)
    Case "frmEASSOC"  'Associations
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_TRADE", "TD_DOC,TD_FILEEXT,TD_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", TD_DOC AS DOC,TD_FILEEXT as FILEEXT, TD_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_TRADE", "TD_DOC,TD_FILEEXT,TD_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", TD_DOC AS DOC,TD_FILEEXT as FILEEXT, TD_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_TRADE WHERE TD_TYPE='" & UCase(glbDocName) & "' AND TD_EMPNBR=" & glbLEE_ID
        End If
        'getSQL = getSQL & " AND TD_DOCKEY= " & glbDocKey & " "
        getSQL = getSQL & " AND TD_CODE = '" & glbAssocCode & "'"
        getSQL = getSQL & " AND TD_BEGINDT = " & Date_SQL(glbBeginDt)
    Case "frmVATTEND"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_ATTENDANCE", "AD_DOC,AD_FILEEXT,AD_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", AD_DOC AS DOC,AD_FILEEXT as FILEEXT, AD_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_ATTENDANCE", "AD_DOC,AD_FILEEXT,AD_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", AD_DOC AS DOC,AD_FILEEXT as FILEEXT, AD_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_ATTENDANCE WHERE AD_TYPE='" & UCase(glbDocName) & "' AND AD_EMPNBR=" & glbLEE_ID
        End If
        'getSQL = getSQL & " AND AD_DOCKEY= " & glbDocKey & " "
        getSQL = getSQL & " AND AD_REASON ='" & glbAttReason & "'"
        getSQL = getSQL & " AND AD_DOA = " & Date_SQL(glbAttDOA)
    Case "frmECOMMENTS" 'Comments
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_COMMENTS", "DO_DOC,DO_FILEEXT,DO_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DO_DOC AS DOC,DO_FILEEXT as FILEEXT, DO_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_COMMENTS", "DO_DOC,DO_FILEEXT,DO_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DO_DOC AS DOC,DO_FILEEXT as FILEEXT, DO_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_COMMENTS WHERE DO_TYPE='" & UCase(glbDocName) & "' AND DO_EMPNBR=" & glbLEE_ID
        End If
        'getSQL = getSQL & " AND DO_COTYPE= '" & glbJob & "'"
        'getSQL = getSQL & " AND DO_EDATE = " & Date_SQL(glbSDate)
        getSQL = getSQL & " AND DO_DOCKEY= " & glbDocKey & " "
    Case "frmEHSAttach" 'Incident
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HEALTH_SAFETY_2", "DE_DOC,DE_FILEEXT,DE_DOCTYPE") '
            getSQL = "SELECT  " & xFList & ",DE_DOC AS DOC, DE_FILEEXT as FILEEXT, DE_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HEALTH_SAFETY_2", "DE_DOC,DE_FILEEXT,DE_DOCTYPE")
            getSQL = "SELECT  " & xFList & ",DE_DOC AS DOC,DE_FILEEXT as FILEEXT, DE_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND DE_CASE= '" & glbJob & "'"
        getSQL = getSQL & " AND DE_DOCNO = '" & glbDocTmp & "'"
        'getSQL = getSQL & " AND DE_FILEEXT is not null"
        'getSQL = getSQL & " AND DE_OCCDATE is not null"
    Case "frmEHSINJURYWF7" 'Injury - WSIB Form 7 document attachment
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HEALTH_SAFETY_CONCERNSWF7", "W7_DOC,W7_FILEEXT,W7_DOCTYPE") '
            getSQL = "SELECT  " & xFList & ",W7_DOC AS DOC, W7_FILEEXT as FILEEXT, W7_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HEALTH_SAFETY_CONCERNSWF7", "W7_DOC,W7_FILEEXT,W7_DOCTYPE")
            getSQL = "SELECT  " & xFList & ",W7_DOC AS DOC,W7_FILEEXT as FILEEXT, W7_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE='" & UCase(glbDocName) & "' AND W7_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND W7_CASE= '" & glbJob & "'"
        getSQL = getSQL & " AND W7_DOCKEY = '" & glbDocKey & "'"
        'getSQL = getSQL & " AND W7_FILEEXT is not null"
        'getSQL = getSQL & " AND W7_OCCDATE is not null"
    Case "frmEInjF7Sections" 'Injury - WSIB Form 7 Sections document attachment
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_OHS_WRITTEN_OFFER", "F7_DOC,F7_FILEEXT,F7_DOCTYPE") '
            getSQL = "SELECT  " & xFList & ",F7_DOC AS DOC, F7_FILEEXT as FILEEXT, F7_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_OHS_WRITTEN_OFFER", "F7_DOC,F7_FILEEXT,F7_DOCTYPE")
            getSQL = "SELECT  " & xFList & ",F7_DOC AS DOC,F7_FILEEXT as FILEEXT, F7_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE='" & UCase(glbDocName) & "' AND F7_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND F7_CASE= '" & glbJob & "'"
        getSQL = getSQL & " AND F7_DOCKEY = '" & glbDocKey & "'"
        'getSQL = getSQL & " AND W7_FILEEXT is not null"
        'getSQL = getSQL & " AND W7_OCCDATE is not null"
    Case "frmECounsel"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_COUNSEL", "DC_DOC,DC_FILEEXT,DC_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DC_DOC AS DOC,DC_FILEEXT as FILEEXT, DC_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_COUNSEL", "DC_DOC,DC_FILEEXT,DC_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DC_DOC AS DOC,DC_FILEEXT as FILEEXT, DC_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_COUNSEL WHERE DC_TYPE='" & UCase(glbDocName) & "' AND DC_EMPNBR=" & glbLEE_ID
        End If
        'getSQL = getSQL & " AND DC_CLTYPE= '" & glbJob & "'"
        'getSQL = getSQL & " AND DC_COUDATE = " & Date_SQL(glbSDate)
        getSQL = getSQL & " AND DC_DOCKEY= " & glbDocKey & " "
    Case "frmEPERFORM" 'performance
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_PERFORM_HISTORY", "DH_DOC,DH_FILEEXT,DH_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DH_DOC AS DOC,DH_FILEEXT as FILEEXT, DH_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_PERFORM_HISTORY", "DH_DOC,DH_FILEEXT,DH_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DH_DOC AS DOC,DH_FILEEXT as FILEEXT, DH_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_PERFORM_HISTORY WHERE DH_TYPE='" & UCase(glbDocName) & "' AND DH_EMPNBR=" & glbLEE_ID
        End If
        'getSQL = getSQL & " AND DH_JOB= '" & glbJob & "'"
        'getSQL = getSQL & " AND DH_PREVDATE = " & Date_SQL(glbSDate)
        If glbDocKey = "" Then
            getSQL = getSQL & " AND DH_DOCKEY= 0"
        Else
            getSQL = getSQL & " AND DH_DOCKEY= " & glbDocKey & " "
        End If

    Case "frmEPOSITION" 'Position
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_JOB_HISTORY", "DJ_DOC,DJ_FILEEXT,DJ_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DJ_DOC AS DOC,DJ_FILEEXT as FILEEXT, DJ_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_JOB_HISTORY", "DJ_DOC,DJ_FILEEXT,DJ_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DJ_DOC AS DOC,DJ_FILEEXT as FILEEXT, DJ_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_JOB_HISTORY WHERE DJ_TYPE='" & UCase(glbDocName) & "' AND DJ_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND DJ_JOB= '" & glbJob & "'"
        getSQL = getSQL & " AND DJ_SDATE = " & Date_SQL(glbSDate)
        
    Case "frmMPOSITIONS" 'Position Master
        xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_JOB", "DB_DOC, DB_FILEEXT,DB_DOCTYPE")
        getSQL = "SELECT  " & xFList & ", DB_DOC AS DOC,DB_FILEEXT as FILEEXT, DB_DOCTYPE as DOCTYPE "
        getSQL = getSQL & " from HRDOC_JOB WHERE DB_TYPE='" & UCase(glbDocName) & "' AND DB_JOB='" & glbPos & "'"
        
    Case "frmEESTATS"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EMP", "RE_DOC,RE_FILEEXT,RE_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", RE_DOC AS DOC,RE_FILEEXT as FILEEXT, RE_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_EMP WHERE RE_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EMP", "RE_DOC,RE_FILEEXT,RE_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", RE_DOC AS DOC,RE_FILEEXT as FILEEXT, RE_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_EMP WHERE RE_TYPE='" & UCase(glbDocName) & "' AND RE_EMPNBR=" & glbLEE_ID
        End If
    Case "frmESEMINARS"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EDSEM", "ES_DOC,ES_FILEEXT,ES_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", ES_DOC AS DOC,ES_FILEEXT as FILEEXT, ES_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_EDSEM WHERE ES_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EDSEM", "ES_DOC,ES_FILEEXT,ES_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", ES_DOC AS DOC,ES_FILEEXT as FILEEXT, ES_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_EDSEM WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND ES_DOCKEY= " & glbDocKey & " "
    
    Case "frmESEMRETEST"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_EDSEM_RETEST", "ES_DOC,ES_FILEEXT,ES_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", ES_DOC AS DOC,ES_FILEEXT as FILEEXT, ES_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_EDSEM_RETEST WHERE ES_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_EDSEM_RETEST", "ES_DOC,ES_FILEEXT,ES_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", ES_DOC AS DOC,ES_FILEEXT as FILEEXT, ES_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_EDSEM_RETEST WHERE ES_TYPE='" & UCase(glbDocName) & "' AND ES_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND ES_DOCKEY= " & glbDocKey & " "
    
    Case "frmFORMALED"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HREDU", "EU_DOC,EU_FILEEXT,EU_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", EU_DOC AS DOC,EU_FILEEXT as FILEEXT, EU_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HREDU", "EU_DOC,EU_FILEEXT,EU_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", EU_DOC AS DOC,EU_FILEEXT as FILEEXT, EU_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_HREDU WHERE EU_TYPE='" & UCase(glbDocName) & "' AND EU_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND EU_DOCKEY= " & glbDocKey & " "
    
    Case "frmEODOLLAR"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_DOLENT", "DE_DOC,DE_FILEEXT,DE_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DE_DOC AS DOC,DE_FILEEXT as FILEEXT, DE_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_DOLENT WHERE DE_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HRDOLENT", "DE_DOC,DE_FILEEXT,DE_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", DE_DOC AS DOC,DE_FILEEXT as FILEEXT, DE_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_HRDOLENT WHERE DE_TYPE='" & UCase(glbDocName) & "' AND DE_EMPNBR=" & glbLEE_ID
        End If
        getSQL = getSQL & " AND DE_DOCKEY= " & glbDocKey & " "
    
    Case "frmPosSkills" 'Position Skills
        xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_JOBSKL", "DS_DOC,DS_FILEEXT,DS_DOCTYPE")
        getSQL = "SELECT  " & xFList & ", DS_DOC AS DOC,DS_FILEEXT as FILEEXT, DS_DOCTYPE as DOCTYPE "
        getSQL = getSQL & " from HRDOC_JOBSKL WHERE DS_TYPE='" & UCase(glbDocName) & "' AND DS_JOB='" & glbPos & "' AND DS_SKILL='" & glbPosSkill & "'"
    
    'Release 8.1
    Case "frmEmpOther"
        If glbtermopen Then
            If glbTERM_Seq = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HREMP_OTHER", "ER_DOC,ER_FILEEXT,ER_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", ER_DOC AS DOC,ER_FILEEXT as FILEEXT, ER_DOCTYPE as DOCTYPE "
            getSQL = getSQL & "  from Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HREMP_OTHER", "ER_DOC,ER_FILEEXT,ER_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", ER_DOC AS DOC,ER_FILEEXT as FILEEXT, ER_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND ER_EMPNBR=" & glbLEE_ID
        End If
    'Release 8.1
    Case "frmETLAY"
        'If glbtermopen Then
        '    If glbTERM_Seq = 0 Then Exit Function
        '
        '    xFList = Get_Fields(gdbAdoIhr001_DOC, "Term_HRDOC_HREMP_OTHER", "ER_DOC,ER_FILEEXT,ER_DOCTYPE")
        '    getSQL = "SELECT  " & xFList & ", ER_DOC AS DOC,ER_FILEEXT as FILEEXT, ER_DOCTYPE as DOCTYPE "
        '    getSQL = getSQL & "  from Term_HRDOC_HREMP_OTHER WHERE ER_TYPE='" & UCase(glbDocName) & "' AND TERM_SEQ = " & glbTERM_Seq
        'Else
            If glbLEE_ID = 0 Then Exit Function
            
            xFList = Get_Fields(gdbAdoIhr001_DOC, "HRDOC_HRSTATUS", "SC_DOC,SC_FILEEXT,SC_DOCTYPE")
            getSQL = "SELECT  " & xFList & ", SC_DOC AS DOC,SC_FILEEXT as FILEEXT, SC_DOCTYPE as DOCTYPE "
            getSQL = getSQL & " from HRDOC_HRSTATUS WHERE SC_TYPE='" & UCase(glbDocName) & "' AND SC_EMPNBR=" & glbLEE_ID
        'End If
        getSQL = getSQL & " AND SC_DOCKEY= " & glbDocKey & " "
        
    End Select
    
End Function

Public Sub DispimgIcon(zForm As Form, zFormName)
Dim SQLQ
Dim rsTemp As New adodb.Recordset

On Error Resume Next

If Not gsAttachment_DB Then Exit Sub

If zFormName = "frmEESTATS" And glbDocName = "Termination" Then
    zForm.lblImport1.Visible = True
ElseIf zFormName <> "frmEmployeeFlags" Then
    zForm.lblImport.Visible = True
End If

SQLQ = getSQL(zFormName)

If zFormName = "frmEmployeeFlags" Then Exit Sub

rsTemp.Open SQLQ, gdbAdoIhr001_DOC, adOpenStatic
If Not rsTemp.EOF Then
    If zFormName = "frmEESTATS" And glbDocName = "Termination" Then
        zForm.imgSec1.Visible = True
        zForm.imgNoSec1.Visible = False
    Else
        zForm.imgSec.Visible = True
        zForm.imgNoSec.Visible = False
    End If
Else
    If zFormName = "frmEESTATS" And glbDocName = "Termination" Then
        zForm.imgSec1.Visible = False
        zForm.imgNoSec1.Visible = True
    Else
        zForm.imgSec.Visible = False
        zForm.imgNoSec.Visible = True
    End If
End If
rsTemp.Close

zForm.cmdImport.Visible = False
Select Case glbDocName
    Case "Resume"
        If gSec_Upd_Basic And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Offer"
        If gSec_Upd_Position And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Comments"
        If gSec_Upd_Comments And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "INCIDENT"
        If gSec_Upd_Health_Safety And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "INJURYWF7"
        If gSec_Upd_Health_Safety And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "INJURYWF7_WRITTENOFR"
        If gSec_Upd_Health_Safety And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Counsel"
        If gSec_Upd_Counselling And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Performance"
        If gSec_Upd_Performance And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Jobdescription"
        If gSec_Upd_Job_Master Then
            zForm.cmdImport.Visible = True
        End If
    Case "EdSem"
        If gSec_Upd_Education_Seminars And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "EdSem_Retest"
        If gSec_Upd_Education_Seminars And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "FormalEdu"
        If gSec_Upd_Formal_Education And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "DollarEnt"
        If gSec_Upd_Other_Entitlements And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Associations"
        If gSec_Upd_Associations Then 'And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Attendance"
        If gSec_Upd_Attendance And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "Termination"
        If gSec_Upd_Terminations And Not glbtermopen Then
            zForm.cmdImport1.Visible = True
        End If
    Case "PositionSkill"
        If gSec_Upd_Job_Skills Then
            zForm.cmdImport.Visible = True
        End If
    'Release 8.1
    Case "OtherInfo"
        If gSec_Upd_OtherInformation And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
    Case "LOA"
        If gSec_Inq_EnterLeave And Not glbtermopen Then
            zForm.cmdImport.Visible = True
        End If
End Select

End Sub

Public Function UserEmailExist()
Dim rsEmail As New adodb.Recordset
    UserEmailExist = True
    rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
    If rsEmail.EOF Then
        UserEmailExist = False
        MsgBox "You have not been set up for email sending.  Please use the Setup->Security->Email Setup menu option to set up your account for emailing.", vbCritical + vbOKOnly, "No Email Setup Found"
    End If
    rsEmail.Close
End Function

Public Function GetComPreferSMTP(xType)
Dim rsSMTP As New adodb.Recordset
Dim SQLQ, xSMTP
    xSMTP = ""  'SMTP Server|SMTP User|SMTP Password|SMTP Port
    
    'Retrieve SMTP Information: HP_SERVER,HP_USERNAME,HP_PASSWORD,HP_PORT
    SQLQ = "SELECT * FROM HRPREFERENCE WHERE HP_FUN_NAME = '" & xType & "' "
    rsSMTP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsSMTP.EOF Then
        If Not IsNull(rsSMTP("HP_SERVER")) Then
            If Len((rsSMTP("HP_SERVER"))) > 0 Then
                xSMTP = rsSMTP("HP_SERVER") & "|"
            End If
        End If
        If Not IsNull(rsSMTP("HP_USERNAME")) Then
            If Len((rsSMTP("HP_USERNAME"))) > 0 Then
                xSMTP = xSMTP & rsSMTP("HP_USERNAME") & "|"
            Else
                xSMTP = xSMTP & "|"
            End If
        Else
            xSMTP = xSMTP & "|"
        End If
        If Not IsNull(rsSMTP("HP_PASSWORD")) Then
            If Len((rsSMTP("HP_PASSWORD"))) > 0 Then
                xSMTP = xSMTP & rsSMTP("HP_PASSWORD") & "|"
            Else
                xSMTP = xSMTP & "|"
            End If
        Else
            xSMTP = xSMTP & "|"
        End If
        If Not IsNull(rsSMTP("HP_PORT")) Then
            If Len((rsSMTP("HP_PORT"))) > 0 Then
                xSMTP = xSMTP & rsSMTP("HP_PORT") & "|"
            Else
                xSMTP = xSMTP & "|"
            End If
        Else
            xSMTP = xSMTP & "|"
        End If
    End If
    rsSMTP.Close
    
    GetComPreferSMTP = xSMTP
End Function

Public Function GetComPreferEmail(xType, Optional xEmpNo, Optional xTERM_Seq)
Dim rsEmail As New adodb.Recordset
Dim SQLQ, xEmail, xEmpEmail
    xEmail = ""
    
    'Ticket #20317 - 'More Emails' option for everyone
    'If glbCompSerial = "S/N - 2382W" And Not IsMissing(xEmpNo) Then  'Samuel Ticket #18090
    If Not IsMissing(xEmpNo) Then  'Samuel Ticket #18090
        If IsMissing(xTERM_Seq) Then
            xEmail = GetComPreferEmailDetails(xType, xEmpNo)
        Else
            xEmail = GetComPreferEmailDetails(xType, xEmpNo, xTERM_Seq)
        End If
        
        ''They don't want this change now - changed their mind
        ''Ticket #27636 - Cascade Canada Ltd.
        ''They want to send email to Employee's Email Address from Status/Dates screen as well
        'If glbCompSerial = "S/N - 2344W" Then
        '    xEmpEmail = ""
        '    xEmpEmail = GetCurEmpEmail
        '    If Len(xEmpEmail) > 0 Then
        '        If Len(xEmail) > 0 Then
        '            xEmail = xEmail & ";" & xEmpEmail
        '        Else
        '            xEmail = xEmpEmail
        '        End If
        '    End If
        'End If
    Else
        SQLQ = "SELECT * FROM HRPREFERENCE WHERE HP_FUN_NAME = '" & xType & "' "
        rsEmail.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmail.EOF Then
            If Not IsNull(rsEmail("HP_EMAIL")) Then
                If Len((rsEmail("HP_EMAIL"))) > 0 Then
                    xEmail = rsEmail("HP_EMAIL")
                End If
            End If
        End If
        rsEmail.Close
    End If
    GetComPreferEmail = xEmail
End Function

Public Function GetReptAuthEmail(xEmpNo, Optional xTERM_Seq) 'Ticket #23453 Franks 03/26/2013
Dim rsEmpJob As New adodb.Recordset
Dim SQLQ
Dim xRAEmpNo
Dim EmpTermSeq
Dim retval
    retval = ""
    If IsMissing(xTERM_Seq) Then EmpTermSeq = 0 Else EmpTermSeq = xTERM_Seq
    If EmpTermSeq = 0 Then
        SQLQ = "SELECT * FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
        SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT * FROM Term_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
        SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & EmpTermSeq & " "
    End If
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpJob.EOF Then
        xRAEmpNo = ""
        If Not IsNull(rsEmpJob("JH_REPTAU")) Then
            If rsEmpJob("JH_REPTAU") > 0 Then
                xRAEmpNo = rsEmpJob("JH_REPTAU")
            End If
        End If
        If Len(xRAEmpNo) = 0 Then 'No RA1 then check RA2
            If Not IsNull(rsEmpJob("JH_REPTAU2")) Then
                If rsEmpJob("JH_REPTAU2") > 0 Then
                    xRAEmpNo = rsEmpJob("JH_REPTAU2")
                End If
            End If
        End If
        If Len(xRAEmpNo) > 0 Then 'found RA emp no
            'get the email from this xRAEmpNo
            retval = GetEmpData(xRAEmpNo, "ED_EMAIL")
        End If
    End If
    GetReptAuthEmail = retval
End Function
Public Function GetComPreferEmailDetUserFlag(xType, xEmpNo, Optional xTERM_Seq) 'Ticket #23453 Franks 03/25/2013
Dim rsEMP As New adodb.Recordset
Dim rsEmail As New adodb.Recordset
Dim SQLQ, xEmail
Dim xEmailType As String
Dim xPosGroup As String
Dim EmpTermSeq
Dim xVE_GRPCD
Dim retval As Boolean

    xEmail = ""
    xEmailType = ""
    retval = False
    Select Case xType
        Case "EMAIL_ONNEWHIRE"
            xEmailType = "New Hire"
        Case "EMAIL_ONPOSITION"
            xEmailType = "Position"
        Case "EMAIL_ONSALARY"
            xEmailType = "Salary"
        Case "EMAIL_ONBENEFIT"
            xEmailType = "Benefits"
        Case "EMAIL_ONTERM"
            xEmailType = "Termination"
        Case "EMAIL_ONREHIRE"
            xEmailType = "Rehire"
        Case "EMAIL_ONLEAVECHANGES"
            xEmailType = "Leave Changes"
        Case "EMAIL_ONPERFORMANCE"
            xEmailType = "Performance"
        Case "EMAIL_ONDEPENDENT"
            xEmailType = "Dependent"
        Case "EMAIL_ONREQUESTAPPROVAL"
            xEmailType = "ESS-Request Approval"
        Case "EMAIL_ONREQUESTSUBMISSION"    'Ticket #27060 - S.U.C.C.E.S.S.
            xEmailType = "ESS-Request Submit"
        Case "EMAIL_ONEMPLOYEEFLAGS"
            xEmailType = "Employee Flags"   'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
    End Select
    
    If Len(xEmailType) > 0 Then
        'Ticket #21444 Franks 02/10/2012 - begin
        If IsMissing(xTERM_Seq) Then EmpTermSeq = 0 Else EmpTermSeq = xTERM_Seq
        xPosGroup = getEmpPosGroup(xEmpNo, EmpTermSeq)
        'Ticket #21444 Franks 02/10/2012 - end
        SQLQ = "SELECT * FROM HRPREEMAIL WHERE VE_TYPE = '" & xEmailType & "' "
        rsEmail.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsEmail.EOF
            If IsMissing(xTERM_Seq) Then
                SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
            Else
                SQLQ = "SELECT ED_EMPNBR FROM TERM_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_Seq & " "
            End If
            If Not IsNull(rsEmail("VE_DIV")) Then
                If Len(rsEmail("VE_DIV")) > 0 Then
                    SQLQ = SQLQ & "AND ED_DIV = '" & rsEmail("VE_DIV") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_DEPT")) Then
                If Len(rsEmail("VE_DEPT")) > 0 Then
                    SQLQ = SQLQ & "AND ED_DEPTNO = '" & rsEmail("VE_DEPT") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_ORG")) Then
                If Len(rsEmail("VE_ORG")) > 0 Then
                    SQLQ = SQLQ & "AND ED_ORG = '" & rsEmail("VE_ORG") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_EMP")) Then
                If Len(rsEmail("VE_EMP")) > 0 Then
                    SQLQ = SQLQ & "AND ED_EMP = '" & rsEmail("VE_EMP") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_PT")) Then
                If Len(rsEmail("VE_PT")) > 0 Then
                    SQLQ = SQLQ & "AND ED_PT = '" & rsEmail("VE_PT") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_LOC")) Then
                If Len(rsEmail("VE_LOC")) > 0 Then
                    SQLQ = SQLQ & "AND ED_LOC = '" & rsEmail("VE_LOC") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_SECTION")) Then
                If Len(rsEmail("VE_SECTION")) > 0 Then
                    SQLQ = SQLQ & "AND ED_SECTION = '" & rsEmail("VE_SECTION") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_ADMINBY")) Then
                If Len(rsEmail("VE_ADMINBY")) > 0 Then
                    SQLQ = SQLQ & "AND ED_ADMINBY = '" & rsEmail("VE_ADMINBY") & "' "
                End If
            End If

            If rsEMP.State <> 0 Then rsEMP.Close
            rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEMP.EOF Then
                If Not IsNull(rsEmail("VE_USER_FLAG")) Then
                    If rsEmail("VE_USER_FLAG") Then
                        retval = True
                    End If
                End If
            End If
            rsEmail.MoveNext
        Loop
        rsEmail.Close
    
    End If
    GetComPreferEmailDetUserFlag = retval

End Function

Public Function GetComPreferEmailDetails(xType, xEmpNo, Optional xTERM_Seq)
Dim rsEMP As New adodb.Recordset
Dim rsEmail As New adodb.Recordset
Dim SQLQ, xEmail
Dim xEmailType As String
Dim xPosGroup As String
Dim EmpTermSeq
Dim xVE_GRPCD

    xEmail = ""
    xEmailType = ""
    Select Case xType
        Case "EMAIL_ONNEWHIRE"
            xEmailType = "New Hire"
        Case "EMAIL_ONPOSITION"
            xEmailType = "Position"
        Case "EMAIL_ONSALARY"
            xEmailType = "Salary"
        Case "EMAIL_ONBENEFIT"
            xEmailType = "Benefits"
        Case "EMAIL_ONTERM"
            xEmailType = "Termination"
        Case "EMAIL_ONREHIRE"
            xEmailType = "Rehire"
        Case "EMAIL_ONLEAVECHANGES"
            xEmailType = "Leave Changes"
        Case "EMAIL_ONPERFORMANCE"
            xEmailType = "Performance"
        Case "EMAIL_ONDEPENDENT"
            xEmailType = "Dependent"
        Case "EMAIL_ONREQUESTAPPROVAL"
            xEmailType = "ESS-Request Approval"
        Case "EMAIL_ONREQUESTSUBMISSION"    'Ticket #27060 - S.U.C.C.E.S.S.
            xEmailType = "ESS-Request Submit"
        Case "EMAIL_ONEMPLOYEEFLAGS"
            xEmailType = "Employee Flags"   'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
        Case "EMAIL_ONHSINCIDENT" 'Ticket #28664 Franks 05/31/2016
            xEmailType = "H&S Incident"
    End Select
    
    If Len(xEmailType) > 0 Then
        'Ticket #21444 Franks 02/10/2012 - begin
        If IsMissing(xTERM_Seq) Then EmpTermSeq = 0 Else EmpTermSeq = xTERM_Seq
        xPosGroup = getEmpPosGroup(xEmpNo, EmpTermSeq)
        'Ticket #21444 Franks 02/10/2012 - end
        SQLQ = "SELECT * FROM HRPREEMAIL WHERE VE_TYPE = '" & xEmailType & "' "
        rsEmail.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsEmail.EOF
            If IsMissing(xTERM_Seq) Then
                SQLQ = "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
            Else
                SQLQ = "SELECT ED_EMPNBR FROM TERM_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
                SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_Seq & " "
            End If
            If Not IsNull(rsEmail("VE_DIV")) Then
                If Len(rsEmail("VE_DIV")) > 0 Then
                    SQLQ = SQLQ & "AND ED_DIV = '" & rsEmail("VE_DIV") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_DEPT")) Then
                If Len(rsEmail("VE_DEPT")) > 0 Then
                    SQLQ = SQLQ & "AND ED_DEPTNO = '" & rsEmail("VE_DEPT") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_ORG")) Then
                If Len(rsEmail("VE_ORG")) > 0 Then
                    SQLQ = SQLQ & "AND ED_ORG = '" & rsEmail("VE_ORG") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_EMP")) Then
                If Len(rsEmail("VE_EMP")) > 0 Then
                    SQLQ = SQLQ & "AND ED_EMP = '" & rsEmail("VE_EMP") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_PT")) Then
                If Len(rsEmail("VE_PT")) > 0 Then
                    SQLQ = SQLQ & "AND ED_PT = '" & rsEmail("VE_PT") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_LOC")) Then
                If Len(rsEmail("VE_LOC")) > 0 Then
                    SQLQ = SQLQ & "AND ED_LOC = '" & rsEmail("VE_LOC") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_SECTION")) Then
                If Len(rsEmail("VE_SECTION")) > 0 Then
                    SQLQ = SQLQ & "AND ED_SECTION = '" & rsEmail("VE_SECTION") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_ADMINBY")) Then
                If Len(rsEmail("VE_ADMINBY")) > 0 Then
                    SQLQ = SQLQ & "AND ED_ADMINBY = '" & rsEmail("VE_ADMINBY") & "' "
                End If
            End If
            If Not IsNull(rsEmail("VE_REGION")) Then 'Ticket #27515 Franks 09/14/2015
                If Len(rsEmail("VE_REGION")) > 0 Then
                    SQLQ = SQLQ & "AND ED_REGION = '" & rsEmail("VE_REGION") & "' "
                End If
            End If

            If rsEMP.State <> 0 Then rsEMP.Close
            rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
            If Not rsEMP.EOF Then
                'Ticket #21444 Franks 02/10/2012 - begin
                xVE_GRPCD = ""
                If Not IsNull(rsEmail("VE_GRPCD")) Then
                    If Len(rsEmail("VE_GRPCD")) > 0 Then
                        xVE_GRPCD = rsEmail("VE_GRPCD")
                    End If
                End If
                If Len(xVE_GRPCD) > 0 And Len(xPosGroup) > 0 Then
                'There is a Position Group on the setup screen
                    If xVE_GRPCD = xPosGroup Then
                    'The emp Position Group = setup Position Group
                        If Not IsNull(rsEmail("VE_EMAIL")) Then
                            xEmail = rsEmail("VE_EMAIL")
                        End If
                    End If
                'Ticket #21444 Franks 02/10/2012 - end
                Else 'No Position Group on the setup screen
                    If Not IsNull(rsEmail("VE_EMAIL")) Then
                        xEmail = rsEmail("VE_EMAIL")
                    End If
                End If
            End If
            rsEmail.MoveNext
        Loop
        rsEmail.Close
    
    End If
    GetComPreferEmailDetails = xEmail

End Function

Public Function getEmpPosGroup(xEmpNo, xEmpTermSeq)
Dim rsEmpJob As New adodb.Recordset
Dim rsHRJOB As New adodb.Recordset
Dim SQLQ As String
Dim xJobCode As String
Dim retval As String
    retval = ""
    xJobCode = ""
    If xEmpTermSeq = 0 Then
        SQLQ = "SELECT JH_EMPNBR, JH_JOB FROM HR_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
        SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT JH_EMPNBR, JH_JOB FROM Term_JOB_HISTORY WHERE NOT (JH_CURRENT = 0) "
        SQLQ = SQLQ & "AND JH_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xEmpTermSeq & " "
    End If
    rsEmpJob.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpJob.EOF Then
        xJobCode = rsEmpJob("JH_JOB")
    End If
    rsEmpJob.Close
    If Len(xJobCode) > 0 Then
        SQLQ = "SELECT JB_CODE,JB_GRPCD FROM HRJOB WHERE JB_CODE = '" & xJobCode & "' "
        rsHRJOB.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsHRJOB.EOF Then
            retval = GetString(rsHRJOB("JB_GRPCD"))
        End If
        rsHRJOB.Close
    End If
    getEmpPosGroup = retval
End Function
Public Function GetEmailBodyForSamuel(xEmpNo, Optional xTERM_Seq)
Dim rsEMP As New adodb.Recordset
Dim SQLQ As String
Dim retval As String
    retval = "Employee # " & xEmpNo & " "
    If IsMissing(xTERM_Seq) Then
        SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME, ED_ADMINBY, ED_SECTION FROM HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
    Else
        SQLQ = "SELECT ED_EMPNBR,ED_SURNAME,ED_FNAME, ED_ADMINBY, ED_SECTION FROM TERM_HREMP WHERE ED_EMPNBR = " & xEmpNo & " "
        SQLQ = SQLQ & "AND TERM_SEQ = " & xTERM_Seq & " "
    End If
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        'retVal = "Employee # " & xEmpNo & " "
        retval = retval & "- " & rsEMP("ED_FNAME") & " " & rsEMP("ED_SURNAME") & " "
        If Not IsNull(rsEMP("ED_ADMINBY")) Then
            If Len(rsEMP("ED_ADMINBY")) > 0 Then
                'retVal = retVal & "in " & lStr("Administered By") & " " & rsEmp("ED_ADMINBY") & " "
                'retVal = retVal & "in " & lStr("Administered By") & " " & GetTablDesc("EDAB", rsEmp("ED_ADMINBY")) & " "
                retval = retval & "in " & lStr("Administered By") & " " & rsEMP("ED_ADMINBY") & " "
            End If
        End If
        If Not IsNull(rsEMP("ED_SECTION")) Then
            If Len(rsEMP("ED_SECTION")) > 0 Then
                ''retVal = retVal & "in " & lStr("Section") & " " & rsEmp("ED_SECTION") & " "
                ''retVal = retVal & "in " & lStr("Section") & " " & GetTablDesc("EDSE", rsEmp("ED_SECTION")) & " "
                'retVal = retVal & "in " & lStr("Section") & " " & rsEmp("ED_SECTION") & " "
                'Ticket #23453 Franks 04/01/2013, Muhammad needs description in the email
                retval = retval & "in " & lStr("Section") & " " & GetTABLDesc("EDSE", rsEMP("ED_SECTION")) & " "
            End If
        End If
    End If
    rsEMP.Close
    GetEmailBodyForSamuel = retval
End Function

Public Function GetCurUserEmail() 'Ticket #19852 Franks 02/14/2011
Dim rsEmail As New adodb.Recordset
Dim retval As String
    retval = ""
    rsEmail.Open "SELECT * FROM HR_EMAIL WHERE EM_USERID='" & Replace(glbUserID, "'", "''") & "'", gdbAdoIhr001
    If Not rsEmail.EOF Then
        retval = rsEmail("EM_ADDRESS")
    End If
    rsEmail.Close
    GetCurUserEmail = retval
End Function

Public Function GetCurEmpEmail()
Dim rsEmail As New adodb.Recordset
Dim SQLQ, xEmail
    xEmail = ""
    If Not glbtermopen Then
        SQLQ = "SELECT ED_EMPNBR, ED_EMAIL FROM HREMP WHERE ED_EMPNBR = " & glbLEE_ID
        rsEmail.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEmail.EOF Then
            If Not IsNull(rsEmail("ED_EMAIL")) Then
                If Len((rsEmail("ED_EMAIL"))) > 0 Then
                    xEmail = rsEmail("ED_EMAIL")
                End If
            End If
        End If
        rsEmail.Close
    Else
        SQLQ = "SELECT ED_EMPNBR, ED_EMAIL FROM Term_HREMP WHERE TERM_SEQ = " & glbTERM_Seq
        rsEmail.Open SQLQ, gdbAdoIhr001X, adOpenStatic
        If Not rsEmail.EOF Then
            If Not IsNull(rsEmail("ED_EMAIL")) Then
                If Len((rsEmail("ED_EMAIL"))) > 0 Then
                    xEmail = rsEmail("ED_EMAIL")
                End If
            End If
        End If
        rsEmail.Close
    End If
    GetCurEmpEmail = xEmail
End Function

' Given a TABL name and key, returns the value, or "" for none found.
Function GetTABLDesc(TablName, TablKey)
    Dim RSTABL As New adodb.Recordset
    Dim SQLQ
    If IsNull(TablKey) Then
        GetTABLDesc = ""
    Else
        SQLQ = "SELECT * FROM HRTABL WHERE TB_NAME = '" & TablName & "' AND TB_KEY = '" & TablKey & "' "
        RSTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If RSTABL.EOF And RSTABL.BOF Then
            GetTABLDesc = ""
        Else
            GetTABLDesc = RSTABL("TB_DESC")
        End If
        RSTABL.Close
        Set RSTABL = Nothing
    End If
End Function

Public Sub AppendResume(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset

    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    'rsPHOTO.Open "select * from HRDOC_EMP WHERE RE_TYPE = '" & UCase(glbDocName) & "' AND RE_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsPHOTO.Open "select * from HRDOC_EMP WHERE RE_TYPE = '" & UCase(glbDocName) & "' AND RE_EMPNBR=" & zEMPNBR, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))
   
    rsPHOTO.AddNew
    rsPHOTO("RE_EMPNBR") = zEMPNBR
    rsPHOTO("RE_COMPNO") = "001"
    rsPHOTO("RE_FILEEXT") = FileExtension
    rsPHOTO("RE_TYPE") = UCase(glbDocName)
    rsPHOTO("RE_LUSER") = glbUserID
    rsPHOTO("RE_LDATE") = Date
    rsPHOTO("RE_LTIME") = Time$
    
    rsPHOTO("RE_DOCTYPE") = DOCType
    rsPHOTO("RE_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!RE_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    
    rsPHOTO.Close
End Sub

Public Sub AppendOtherInfo(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset

    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    'rsPHOTO.Open "select * from HRDOC_HREMP_OTHER WHERE ER_TYPE = '" & UCase(glbDocName) & "' AND ER_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    rsPHOTO.Open "select * from HRDOC_HREMP_OTHER WHERE ER_TYPE = '" & UCase(glbDocName) & "' AND ER_EMPNBR=" & zEMPNBR, gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))
   
    rsPHOTO.AddNew
    rsPHOTO("ER_EMPNBR") = zEMPNBR
    rsPHOTO("ER_COMPNO") = "001"
    rsPHOTO("ER_FILEEXT") = FileExtension
    rsPHOTO("ER_TYPE") = UCase(glbDocName)
    rsPHOTO("ER_LUSER") = glbUserID
    rsPHOTO("ER_LDATE") = Date
    rsPHOTO("ER_LTIME") = Time$
    
    rsPHOTO("ER_DOCTYPE") = DOCType
    rsPHOTO("ER_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!ER_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    
    rsPHOTO.Close
End Sub

Public Sub AppendOffer(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset

    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    rsPHOTO.Open "select * from HRDOC_JOB_HISTORY WHERE DJ_TYPE = '" & UCase(glbDocName) & "' AND DJ_EMPNBR=" & zEMPNBR & " AND DJ_JOB='" & glbJob & "' AND DJ_SDATE =" & Date_SQL(glbSDate), gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("DJ_EMPNBR") = zEMPNBR
    rsPHOTO("DJ_SDATE") = glbSDate
    rsPHOTO("DJ_JOB") = glbJob
    rsPHOTO("DJ_COMPNO") = "001"
    rsPHOTO("DJ_FILEEXT") = FileExtension
    rsPHOTO("DJ_TYPE") = UCase(glbDocName)
    rsPHOTO("DJ_LUSER") = glbUserID
    rsPHOTO("DJ_LDATE") = Date
    rsPHOTO("DJ_LTIME") = Time$
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DJ_DOCTYPE") = DOCType
    rsPHOTO("DJ_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DJ_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    
    rsPHOTO.Close
End Sub

Public Sub AppendJobdescription(zJobCode, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset

    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If zJobCode = "" Then Exit Sub
    
    rsPHOTO.Open "select * from HRDOC_JOB WHERE DB_TYPE = '" & UCase(glbDocName) & "' and DB_JOB='" & zJobCode & "' ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("DB_JOB") = zJobCode
    rsPHOTO("DB_COMPNO") = "001"
    rsPHOTO("DB_FILEEXT") = FileExtension
    rsPHOTO("DB_TYPE") = UCase(glbDocName)
    rsPHOTO("DB_LUSER") = glbUserID
    rsPHOTO("DB_LDATE") = Date
    rsPHOTO("DB_LTIME") = Time$
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DB_DOCTYPE") = DOCType
    rsPHOTO("DB_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DB_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
End Sub

Public Sub AppendPositionSkill(zJobCode, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset

    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If zJobCode = "" Then Exit Sub
    
    rsPHOTO.Open "select * from HRDOC_JOBSKL WHERE DS_TYPE = '" & UCase(glbDocName) & "' and DS_JOB='" & zJobCode & "' and DS_SKILL = '" & glbPosSkill & "' ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("DS_JOB") = zJobCode
    rsPHOTO("DS_SKILL") = glbPosSkill
    rsPHOTO("DS_COMPNO") = "001"
    rsPHOTO("DS_FILEEXT") = FileExtension
    rsPHOTO("DS_TYPE") = UCase(glbDocName)
    rsPHOTO("DS_LUSER") = glbUserID
    rsPHOTO("DS_LDATE") = Date
    rsPHOTO("DS_LTIME") = Time$
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DS_DOCTYPE") = DOCType
    rsPHOTO("DS_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DS_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
End Sub

Public Sub AppendDollarEnt(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    rsPHOTO.Open "select * from HRDOC_HRDOLENT WHERE DE_TYPE = '" & UCase(glbDocName) & "' AND DE_EMPNBR=" & zEMPNBR & " AND DE_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("DE_EMPNBR") = zEMPNBR
    rsPHOTO("DE_COMPNO") = "001"
    rsPHOTO("DE_FILEEXT") = FileExtension
    rsPHOTO("DE_TYPE") = UCase(glbDocName)
    rsPHOTO("DE_LUSER") = glbUserID
    rsPHOTO("DE_LDATE") = Date
    rsPHOTO("DE_LTIME") = Time$
    rsPHOTO("DE_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DE_DOCTYPE") = DOCType
    rsPHOTO("DE_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DE_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the CL_DOCKEY with CL_ID
    SQLQ = "UPDATE HRDOLENT SET DE_DOCKEY = DE_ENTITLE_ID WHERE DE_ENTITLE_ID=" & glbDocKey
    gdbAdoIhr001.Execute SQLQ
    
End Sub

Public Sub AppendFormalEdu(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    If glbtermopen Then
        rsPHOTO.Open "select * from TERM_HRDOC_HREDU WHERE EU_TYPE = '" & UCase(glbDocName) & "' AND EU_EMPNBR=" & zEMPNBR & " AND TERM_SEQ = " & glbTERM_Seq & " AND EU_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    Else
        'rsPHOTO.Open "select * from HRDOC_COUNSEL WHERE DC_TYPE = '" & UCase(glbDocName) & "' AND DC_EMPNBR=" & zEMPNBR & " AND DC_CLTYPE='" & glbJob & "' AND DC_COUDATE =" & Date_SQL(glbSDate), gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        rsPHOTO.Open "select * from HRDOC_HREDU WHERE EU_TYPE = '" & UCase(glbDocName) & "' AND EU_EMPNBR=" & zEMPNBR & " AND EU_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    End If
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("EU_EMPNBR") = zEMPNBR
    rsPHOTO("EU_COMPNO") = "001"
    rsPHOTO("EU_FILEEXT") = FileExtension
    rsPHOTO("EU_TYPE") = UCase(glbDocName)
    rsPHOTO("EU_LUSER") = glbUserID
    rsPHOTO("EU_LDATE") = Date
    rsPHOTO("EU_LTIME") = Time$
    rsPHOTO("EU_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("EU_DOCTYPE") = DOCType
    rsPHOTO("EU_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!EU_DOC.AppendChunk byteChunk
    Close FileNumber
    
    If glbtermopen Then
        rsPHOTO("TERM_SEQ") = glbTERM_Seq
    End If
    
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the CL_DOCKEY with CL_ID
    If glbtermopen Then
        SQLQ = "UPDATE TERM_EDU SET EU_DOCKEY = EU_ID WHERE EU_ID=" & glbDocKey
        gdbAdoIhr001X.Execute SQLQ
    Else
        SQLQ = "UPDATE HREDU SET EU_DOCKEY = EU_ID WHERE EU_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    End If

End Sub

Public Sub AppendEdSem(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    'rsPHOTO.Open "select * from HRDOC_COUNSEL WHERE DC_TYPE = '" & UCase(glbDocName) & "' AND DC_EMPNBR=" & zEMPNBR & " AND DC_CLTYPE='" & glbJob & "' AND DC_COUDATE =" & Date_SQL(glbSDate), gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    rsPHOTO.Open "select * from HRDOC_EDSEM WHERE ES_TYPE = '" & UCase(glbDocName) & "' AND ES_EMPNBR=" & zEMPNBR & " AND ES_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("ES_EMPNBR") = zEMPNBR
    rsPHOTO("ES_COMPNO") = "001"
    rsPHOTO("ES_FILEEXT") = FileExtension
    rsPHOTO("ES_TYPE") = UCase(glbDocName)
    rsPHOTO("ES_LUSER") = glbUserID
    rsPHOTO("ES_LDATE") = Date
    rsPHOTO("ES_LTIME") = Time$
    rsPHOTO("ES_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("ES_DOCTYPE") = DOCType
    rsPHOTO("ES_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!ES_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the CL_DOCKEY with CL_ID
    SQLQ = "UPDATE HREDSEM SET ES_DOCKEY = ES_ID WHERE ES_ID=" & glbDocKey
    gdbAdoIhr001.Execute SQLQ
    
End Sub

Public Sub AppendEdSem_Retest(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    rsPHOTO.Open "select * from HRDOC_EDSEM_RETEST WHERE ES_TYPE = '" & UCase(glbDocName) & "' AND ES_EMPNBR=" & zEMPNBR & " AND ES_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("ES_EMPNBR") = zEMPNBR
    rsPHOTO("ES_COMPNO") = "001"
    rsPHOTO("ES_FILEEXT") = FileExtension
    rsPHOTO("ES_TYPE") = UCase(glbDocName)
    rsPHOTO("ES_LUSER") = glbUserID
    rsPHOTO("ES_LDATE") = Date
    rsPHOTO("ES_LTIME") = Time$
    rsPHOTO("ES_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("ES_DOCTYPE") = DOCType
    rsPHOTO("ES_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!ES_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the CL_DOCKEY with CL_ID
    SQLQ = "UPDATE HREDSEM_RETEST SET ES_DOCKEY = ES_ID WHERE ES_ID=" & glbDocKey
    gdbAdoIhr001.Execute SQLQ
    
End Sub

Public Sub AppendCounsel(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    If glbtermopen Then
        rsPHOTO.Open "select * from TERM_HRDOC_COUNSEL WHERE DC_TYPE = '" & UCase(glbDocName) & "' AND DC_EMPNBR=" & zEMPNBR & " AND TERM_SEQ = " & glbTERM_Seq & " AND DC_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    Else
        'rsPHOTO.Open "select * from HRDOC_COUNSEL WHERE DC_TYPE = '" & UCase(glbDocName) & "' AND DC_EMPNBR=" & zEMPNBR & " AND DC_CLTYPE='" & glbJob & "' AND DC_COUDATE =" & Date_SQL(glbSDate), gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
        rsPHOTO.Open "select * from HRDOC_COUNSEL WHERE DC_TYPE = '" & UCase(glbDocName) & "' AND DC_EMPNBR=" & zEMPNBR & " AND DC_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    End If
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("DC_EMPNBR") = zEMPNBR
    If Not IsNull(glbCounselDate) And IsDate(glbCounselDate) Then rsPHOTO("DC_COUDATE") = glbCounselDate
    rsPHOTO("DC_CLTYPE") = glbCounselType
    rsPHOTO("DC_COMPNO") = "001"
    rsPHOTO("DC_FILEEXT") = FileExtension
    rsPHOTO("DC_TYPE") = UCase(glbDocName)
    rsPHOTO("DC_LUSER") = glbUserID
    rsPHOTO("DC_LDATE") = Date
    rsPHOTO("DC_LTIME") = Time$
    rsPHOTO("DC_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DC_DOCTYPE") = DOCType
    rsPHOTO("DC_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DC_DOC.AppendChunk byteChunk
    Close FileNumber
    
    If glbtermopen Then
        rsPHOTO("TERM_SEQ") = glbTERM_Seq
    End If
    
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the CL_DOCKEY with CL_ID
    If glbtermopen Then
        SQLQ = "UPDATE TERM_HR_COUNSEL SET CL_DOCKEY = CL_ID WHERE CL_ID=" & glbDocKey
        gdbAdoIhr001X.Execute SQLQ
    Else
        SQLQ = "UPDATE HR_COUNSEL SET CL_DOCKEY = CL_ID WHERE CL_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub

Public Sub AppendComments(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    If glbtermopen Then
        rsPHOTO.Open "select * from TERM_HRDOC_COMMENTS WHERE DO_TYPE = '" & UCase(glbDocName) & "' AND DO_EMPNBR=" & zEMPNBR & " AND TERM_SEQ = " & glbTERM_Seq & " AND DO_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    Else
        rsPHOTO.Open "select * from HRDOC_COMMENTS WHERE DO_TYPE = '" & UCase(glbDocName) & "' AND DO_EMPNBR=" & zEMPNBR & " AND DO_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    End If
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("DO_EMPNBR") = zEMPNBR
    If Not IsNull(glbCommentDate) And IsDate(glbCommentDate) Then rsPHOTO("DO_EDATE") = glbCommentDate
    rsPHOTO("DO_COTYPE") = glbCommentType
    rsPHOTO("DO_COMPNO") = "001"
    rsPHOTO("DO_FILEEXT") = FileExtension
    rsPHOTO("DO_TYPE") = UCase(glbDocName)
    rsPHOTO("DO_LUSER") = glbUserID
    rsPHOTO("DO_LDATE") = Date
    rsPHOTO("DO_LTIME") = Time$
    rsPHOTO("DO_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DO_DOCTYPE") = DOCType
    rsPHOTO("DO_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DO_DOC.AppendChunk byteChunk
    Close FileNumber
    
    If glbtermopen Then
        rsPHOTO("TERM_SEQ") = glbTERM_Seq
    End If
    
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the CO_DOCKEY with CO_COMMENT_ID
    If glbtermopen Then
        SQLQ = "UPDATE TERM_COMMENTS SET CO_DOCKEY = CO_COMMENT_ID WHERE CO_COMMENT_ID=" & glbDocKey
        gdbAdoIhr001X.Execute SQLQ
    Else
        SQLQ = "UPDATE HR_COMMENTS SET CO_DOCKEY = CO_COMMENT_ID WHERE CO_COMMENT_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub

Public Sub AppendIncident(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset

    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    'rsPHOTO.Open "select * from HRDOC_HEALTH_SAFETY WHERE DE_TYPE = '" & UCase(glbDocName) & "' AND DE_EMPNBR=" & zEMPNBR & " AND DE_CASE='" & glbJob & "' AND DE_DOCNO ='" & frmEHSAttach.txtDocNum & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    rsPHOTO.Open "select * from HRDOC_HEALTH_SAFETY_2 WHERE DE_TYPE = '" & UCase(glbDocName) & "' AND DE_EMPNBR=" & zEMPNBR & " AND DE_CASE='" & glbJob & "' AND DE_DOCNO ='" & glbDocTmp & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    'glbDocKey
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
        
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))
    
    rsPHOTO.AddNew
    rsPHOTO("DE_EMPNBR") = zEMPNBR
    'rsPHOTO("DE_COMPNO") = "001"
    rsPHOTO("DE_TYPE") = UCase(glbDocName)
    rsPHOTO("DE_CASE") = glbJob
    rsPHOTO("DE_DOCNO") = glbDocTmp ' 'frmEHSAttach.txtDocNum
    rsPHOTO("DE_FILEEXT") = FileExtension
    rsPHOTO("DE_LDATE") = Date
    rsPHOTO("DE_LTIME") = Time$ 'glbSDate
    'rsPHOTO("DE_OCCDATE") = Date
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DE_DOCTYPE") = DOCType
    rsPHOTO("DE_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DE_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
End Sub

Public Sub AppendInjuryWF7(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ  As String
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    'rsPHOTO.Open "select * from HRDOC_HEALTH_SAFETY WHERE DE_TYPE = '" & UCase(glbDocName) & "' AND DE_EMPNBR=" & zEMPNBR & " AND DE_CASE='" & glbJob & "' AND DE_DOCNO ='" & frmEHSAttach.txtDocNum & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    rsPHOTO.Open "select * from HRDOC_HEALTH_SAFETY_CONCERNSWF7 WHERE W7_TYPE = '" & UCase(glbDocName) & "' AND W7_EMPNBR=" & zEMPNBR & " AND W7_CASE='" & glbJob & "' AND W7_DOCKEY ='" & glbDocKey & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    'glbDocKey
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))
    
    rsPHOTO.AddNew
    rsPHOTO("W7_EMPNBR") = zEMPNBR
    rsPHOTO("W7_COMPNO") = "001"
    rsPHOTO("W7_TYPE") = UCase(glbDocName)
    rsPHOTO("W7_CASE") = glbJob
    rsPHOTO("W7_DOCKEY") = glbDocKey ' 'frmEHSAttach.txtDocNum
    rsPHOTO("W7_FILEEXT") = FileExtension
    rsPHOTO("W7_LDATE") = Date
    rsPHOTO("W7_LTIME") = Time$ 'glbSDate
    rsPHOTO("W7_LUSER") = glbUserID
    'rsPHOTO("DE_OCCDATE") = Date
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("W7_DOCTYPE") = DOCType
    rsPHOTO("w7_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!W7_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    If glbtermopen Then
        SQLQ = "UPDATE TERM_HR_OCC_HEALTH_SAFETY SET EC_DOCKEY = EC_ID WHERE EC_ID=" & glbDocKey
        gdbAdoIhr001X.Execute SQLQ
    Else
        SQLQ = "UPDATE HR_OCC_HEALTH_SAFETY SET EC_DOCKEY = EC_ID WHERE EC_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub

Public Sub AppendInjuryWrittenOfferWF7(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ  As String
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    'rsPHOTO.Open "select * from HRDOC_HEALTH_SAFETY WHERE DE_TYPE = '" & UCase(glbDocName) & "' AND DE_EMPNBR=" & zEMPNBR & " AND DE_CASE='" & glbJob & "' AND DE_DOCNO ='" & frmEHSAttach.txtDocNum & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    rsPHOTO.Open "select * from HRDOC_OHS_WRITTEN_OFFER WHERE F7_TYPE = '" & UCase(glbDocName) & "' AND F7_EMPNBR=" & zEMPNBR & " AND F7_CASE='" & glbJob & "' AND F7_DOCKEY ='" & glbDocKey & "'", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    'glbDocKey
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))
    
    rsPHOTO.AddNew
    rsPHOTO("F7_EMPNBR") = zEMPNBR
    rsPHOTO("F7_COMPNO") = "001"
    rsPHOTO("F7_TYPE") = UCase(glbDocName)
    rsPHOTO("F7_CASE") = glbJob
    rsPHOTO("F7_DOCKEY") = glbDocKey ' 'frmEHSAttach.txtDocNum
    rsPHOTO("F7_FILEEXT") = FileExtension
    rsPHOTO("F7_LDATE") = Date
    rsPHOTO("F7_LTIME") = Time$ 'glbSDate
    rsPHOTO("F7_LUSER") = glbUserID
    'rsPHOTO("DE_OCCDATE") = Date
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("F7_DOCTYPE") = DOCType
    rsPHOTO("F7_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!F7_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    If glbtermopen Then
        SQLQ = "UPDATE TERM_OHS_FORM7_SECTIONS SET F7_DOCKEY = F7_ID WHERE F7_ID=" & glbDocKey
        gdbAdoIhr001X.Execute SQLQ
    Else
        SQLQ = "UPDATE HR_OHS_FORM7_SECTIONS SET F7_DOCKEY = F7_ID WHERE F7_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub

Public Sub AppendPerformance(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    'rsPHOTO.Open "select * from HRDOC_PERFORM_HISTORY WHERE DH_TYPE = '" & UCase(glbDocName) & "' AND DH_EMPNBR=" & zEMPNBR & " AND DH_JOB='" & glbJob & "' AND DH_PREVDATE =" & Date_SQL(glbSDate), gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    rsPHOTO.Open "select * from HRDOC_PERFORM_HISTORY WHERE DH_TYPE = '" & UCase(glbDocName) & "' AND DH_EMPNBR=" & zEMPNBR & " AND DH_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("DH_EMPNBR") = zEMPNBR
    rsPHOTO("DH_PREVDATE") = glbSDate
    rsPHOTO("DH_JOB") = glbJob
    rsPHOTO("DH_COMPNO") = "001"
    rsPHOTO("DH_FILEEXT") = FileExtension
    rsPHOTO("DH_TYPE") = UCase(glbDocName)
    rsPHOTO("DH_LUSER") = glbUserID
    rsPHOTO("DH_LDATE") = Date
    rsPHOTO("DH_LTIME") = Time$
    rsPHOTO("DH_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("DH_DOCTYPE") = DOCType
    rsPHOTO("DH_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!DH_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the PH_DOCKEY with PH_ID
    SQLQ = "UPDATE HR_PERFORM_HISTORY SET PH_DOCKEY = PH_ID WHERE PH_ID=" & glbDocKey
    gdbAdoIhr001.Execute SQLQ
    
End Sub

Public Sub AppendAssociations(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    If glbtermopen Then
        rsPHOTO.Open "select * from TERM_HRDOC_TRADE WHERE TD_TYPE = '" & UCase(glbDocName) & "' AND TD_EMPNBR=" & zEMPNBR & " AND TERM_SEQ = " & glbTERM_Seq & " AND TD_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    Else
        rsPHOTO.Open "select * from HRDOC_TRADE WHERE TD_TYPE = '" & UCase(glbDocName) & "' AND TD_EMPNBR=" & zEMPNBR & " AND TD_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    End If
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("TD_EMPNBR") = zEMPNBR
    rsPHOTO("TD_CODE") = glbAssocCode
    rsPHOTO("TD_BEGINDT") = glbBeginDt
    rsPHOTO("TD_COMPNO") = "001"
    rsPHOTO("TD_FILEEXT") = FileExtension
    rsPHOTO("TD_TYPE") = UCase(glbDocName)
    rsPHOTO("TD_LUSER") = glbUserID
    rsPHOTO("TD_LDATE") = Date
    rsPHOTO("TD_LTIME") = Time$
    rsPHOTO("TD_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("TD_DOCTYPE") = DOCType
    rsPHOTO("TD_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!TD_DOC.AppendChunk byteChunk
    Close FileNumber
    
    If glbtermopen Then
        rsPHOTO("TERM_SEQ") = glbTERM_Seq
    End If
    
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the TD_DOCKEY with TD_ID
    If glbtermopen Then
        SQLQ = "UPDATE TERM_TRADE SET TD_DOCKEY = TD_ID WHERE TD_ID=" & glbDocKey
        gdbAdoIhr001X.Execute SQLQ
    Else
        SQLQ = "UPDATE HRTRADE SET TD_DOCKEY = TD_ID WHERE TD_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub

Public Sub AppendAttendance(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    If glbtermopen Then
        rsPHOTO.Open "select * from TERM_HRDOC_ATTENDANCE WHERE AD_TYPE = '" & UCase(glbDocName) & "' AND AD_EMPNBR=" & zEMPNBR & " AND TERM_SEQ = " & glbTERM_Seq & " AND AD_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    Else
        rsPHOTO.Open "select * from HRDOC_ATTENDANCE WHERE AD_TYPE = '" & UCase(glbDocName) & "' AND AD_EMPNBR=" & zEMPNBR & " AND AD_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    End If
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("AD_EMPNBR") = zEMPNBR
    rsPHOTO("AD_REASON") = glbAttReason
    rsPHOTO("AD_DOA") = glbAttDOA
    rsPHOTO("AD_COMPNO") = "001"
    rsPHOTO("AD_FILEEXT") = FileExtension
    rsPHOTO("AD_TYPE") = UCase(glbDocName)
    rsPHOTO("AD_LUSER") = glbUserID
    rsPHOTO("AD_LDATE") = Date
    rsPHOTO("AD_LTIME") = Time$
    rsPHOTO("AD_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("AD_DOCTYPE") = DOCType
    rsPHOTO("AD_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!AD_DOC.AppendChunk byteChunk
    Close FileNumber
    
    If glbtermopen Then
        rsPHOTO("TERM_SEQ") = glbTERM_Seq
    End If
    
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the AD_DOCKEY with AD_ATT_ID
    If glbtermopen Then
        SQLQ = "UPDATE TERM_ATTENDANCE SET AD_DOCKEY = AD_ATT_ID WHERE AD_ATT_ID=" & glbDocKey
        gdbAdoIhr001X.Execute SQLQ
    Else
        SQLQ = "UPDATE HR_ATTENDANCE SET AD_DOCKEY = AD_ATT_ID WHERE AD_ATT_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    End If
    
End Sub

Public Sub AppendEmployeeFlag(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    rsPHOTO.Open "select * from HRDOC_EMP_FLAGS WHERE EF_FLAG = " & glbEmpFlagNo & " AND EF_TYPE = '" & UCase(glbDocName) & "' AND EF_EMPNBR=" & zEMPNBR & " AND EF_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("EF_EMPNBR") = zEMPNBR
    rsPHOTO("EF_COMPNO") = "001"
    rsPHOTO("EF_FLAG") = glbEmpFlagNo
    rsPHOTO("EF_FLAGDTE") = glbEmpFlagDate
    rsPHOTO("EF_FILEEXT") = FileExtension
    rsPHOTO("EF_TYPE") = UCase(glbDocName)
    rsPHOTO("EF_LUSER") = glbUserID
    rsPHOTO("EF_LDATE") = Date
    rsPHOTO("EF_LTIME") = Time$
    rsPHOTO("EF_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("EF_DOCTYPE") = DOCType
    rsPHOTO("EF_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!EF_DOC.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the EF_DOCKEY with EF_ID
    SQLQ = "UPDATE HREMP_FLAGS SET EF_DOCKEY = EF_ID WHERE EF_ID=" & glbDocKey
    gdbAdoIhr001.Execute SQLQ
    
End Sub

Public Sub AppendLOA(zEMPNBR, FileName As String, FileExtension As String, DOCType As String, UserDesc As String)
    Dim rsPHOTO As New adodb.Recordset
    Dim SQLQ
    Dim byteChunk() As Byte
    Dim x, xChr
    Dim FileNumber As Integer
    
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    
    'If glbtermopen Then
    '    rsPHOTO.Open "select * from TERM_HRDOC_HRSTATUS WHERE SC_TYPE = '" & UCase(glbDocName) & "' AND SC_EMPNBR=" & zEMPNBR & " AND TERM_SEQ = " & glbTERM_Seq & " AND SC_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    'Else
        rsPHOTO.Open "select * from HRDOC_HRSTATUS WHERE SC_TYPE = '" & UCase(glbDocName) & "' AND SC_EMPNBR=" & zEMPNBR & " AND SC_DOCKEY =" & glbDocKey & " ", gdbAdoIhr001_DOC, adOpenKeyset, adLockOptimistic
    'End If
    If Not rsPHOTO.EOF Then
        rsPHOTO.Delete
    End If
    
    glbUPDTCNT = glbUPDTCNT + 1
    
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))

    rsPHOTO.AddNew
    rsPHOTO("SC_EMPNBR") = zEMPNBR
    'rsPHOTO("SC_FDATE") = glbSDate
    'rsPHOTO("SC_STYPE") = glbJob
    rsPHOTO("SC_COMPNO") = "001"
    rsPHOTO("SC_FILEEXT") = FileExtension
    rsPHOTO("SC_TYPE") = UCase(glbDocName)
    rsPHOTO("SC_LUSER") = glbUserID
    rsPHOTO("SC_LDATE") = Date
    rsPHOTO("SC_LTIME") = Time$
    rsPHOTO("SC_DOCKEY") = glbDocKey
    
    '8.0 - Ticket #22682 - Add Document Type and User Description
    rsPHOTO("SC_DOCTYPE") = DOCType
    rsPHOTO("SC_USRDESC") = UserDesc
    
    Get FileNumber, , byteChunk
    rsPHOTO!SC_DOC.AppendChunk byteChunk
    Close FileNumber
    
    'If glbtermopen Then
    '    rsPHOTO("TERM_SEQ") = glbTERM_Seq
    'End If
    
    If glbSQL Or glbOracle Then rsPHOTO.Update
    rsPHOTO.Close
    
    'Update the SC_DOCKEY with SC_ID
    'If glbtermopen Then
    '    SQLQ = "UPDATE TERM_HRSTATUS SET SC_DOCKEY = SC_ID WHERE SC_ID=" & glbDocKey
    '    gdbAdoIhr001X.Execute SQLQ
    'Else
        SQLQ = "UPDATE HRSTATUS SET SC_DOCKEY = SC_ID WHERE SC_ID=" & glbDocKey
        gdbAdoIhr001.Execute SQLQ
    'End If
    
End Sub

Public Function GetFileExtension(xFileName)
Dim I As Integer
Dim xRetVal As String
    For I = 1 To Len(xFileName)
        If Mid(xFileName, Len(xFileName) - I + 1, 1) = "." Then
            xRetVal = Mid(xFileName, Len(xFileName) - I + 2, I - 1)
            GoTo End_line
        End If
    Next
End_line:
    GetFileExtension = xRetVal
End Function

Public Sub AttachmentAdd(xEmpnbr, xFileName As String, xDocType As String, xDocDesc As String)
Dim rsEMP As New adodb.Recordset
Dim ImpFlag As Boolean
Dim xFileExtension As String

    If xEmpnbr <> 0 Then
        If glbtermopen Then
            rsEMP.Open "SELECT ED_EMPNBR FROM TERM_HREMP WHERE ED_EMPNBR=" & xEmpnbr & " AND TERM_SEQ = " & glbTERM_Seq & " AND " & glbSeleDeptUn, gdbAdoIhr001X, adOpenStatic
        Else
            rsEMP.Open "SELECT ED_EMPNBR FROM HREMP WHERE ED_EMPNBR=" & xEmpnbr & " AND " & glbSeleDeptUn, gdbAdoIhr001, adOpenStatic
        End If
        If Not rsEMP.EOF Then
            ImpFlag = True
            'File1.selected(X) = False
        End If
        rsEMP.Close
    End If
    
    If glbLinamar And glbLinHS Then 'Ticket #12401 Facility Health Safety
        ImpFlag = True
    End If
    
    'Ticket #16304
    'xFileExtension = Mid(xFileName, InStr(xFileName, ".") + 1, Len(xFileName) - InStr(xFileName, ".")) 'Ticket #15822
    xFileExtension = GetFileExtension(xFileName)
    
    If ImpFlag Then
        Select Case glbDocName
        Case "Resume"
            Call AppendResume(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "Offer"
            Call AppendOffer(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "INCIDENT"
            Call AppendIncident(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "INJURYWF7"
            Call AppendInjuryWF7(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "INJURYWF7_WRITTENOFR"
            Call AppendInjuryWrittenOfferWF7(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "Comments"
            Call AppendComments(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "Counsel"
            Call AppendCounsel(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "Performance"
            Call AppendPerformance(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "EdSem"
            Call AppendEdSem(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "EdSem_Retest"
            Call AppendEdSem_Retest(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "FormalEdu"
            Call AppendFormalEdu(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "DollarEnt"
            Call AppendDollarEnt(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "Associations"
            Call AppendAssociations(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "Attendance"
            Call AppendAttendance(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "Termination"
            Call AppendResume(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "EmployeeFlag"
            Call AppendEmployeeFlag(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        'Release 8.1
        Case "OtherInfo"
            Call AppendOtherInfo(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        Case "LOA"
            Call AppendLOA(xEmpnbr, xFileName, xFileExtension, xDocType, xDocDesc)
        End Select
    End If
    
    'Jobdescription
    If glbPos <> "" And glbDocName = "Jobdescription" Then
        Call AppendJobdescription(glbPos, xFileName, xFileExtension, xDocType, xDocDesc)
    ElseIf glbPos <> "" And glbDocName = "PositionSkill" Then
        Call AppendPositionSkill(glbPos, xFileName, xFileExtension, xDocType, xDocDesc)
    End If
End Sub

Public Function getCrsCodeMasterFlag()
Dim rsFlag As New adodb.Recordset
Dim SQLQ As String
    getCrsCodeMasterFlag = False
    SQLQ = "SELECT * FROM HR_COURSECODE_MASTER"
    rsFlag.CacheSize = 1
    rsFlag.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    If Not rsFlag.EOF Then
        getCrsCodeMasterFlag = True
    End If
    rsFlag.Close
End Function

Public Function GetCrsCodeDesc(TablKey)
    Dim RSTABL As New adodb.Recordset
    Dim SQLQ
    SQLQ = "SELECT TB_DESC FROM HRTABL WHERE TB_NAME = 'ESCD' AND TB_KEY = '" & TablKey & "' "
    RSTABL.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If RSTABL.EOF And RSTABL.BOF Then
        GetCrsCodeDesc = ""
    Else
        GetCrsCodeDesc = RSTABL("TB_DESC")
    End If
    RSTABL.Close
End Function

Public Sub UpdPswExpireDatac(xUserID, xOldPsw)
    Dim rsUserSecure As New adodb.Recordset
    Dim SQLQ As String
    Dim I As Integer
    
        SQLQ = "SELECT * FROM HR_SECURE_BASIC "
        SQLQ = SQLQ & "Where (USERID = '" & Replace(xUserID, "'", "''") & "')"
        rsUserSecure.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not (rsUserSecure.BOF And rsUserSecure.EOF) Then
            If IsNull(rsUserSecure("PS_EXPIR_DAYS")) Then
                rsUserSecure("PS_EXPIR_DAYS") = 0
            End If
            rsUserSecure("PS_CHGDATE") = Date
            rsUserSecure("PS_EXPIR_DATE") = DateAdd("D", rsUserSecure("PS_EXPIR_DAYS"), Date)
            rsUserSecure("PS_OLDPW3") = rsUserSecure("PS_OLDPW2")
            rsUserSecure("PS_OLDPW2") = rsUserSecure("PS_OLDPW")
            rsUserSecure("PS_OLDPW") = xOldPsw
            rsUserSecure.Update
        End If
        rsUserSecure.Close
End Sub
Public Function EmpNoInTerm(xEmpNo) As Boolean
    Dim rsEmpTerm As New adodb.Recordset
    Dim SQLQ As String
    
    EmpNoInTerm = False
    SQLQ = "SELECT ED_EMPNBR FROM Term_HREMP WHERE ED_EMPNBR = " & xEmpNo
    rsEmpTerm.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEmpTerm.EOF Then
        EmpNoInTerm = True
    End If
    rsEmpTerm.Close
    
End Function

Public Function isTransferGP(Product_Info, xFunction)
Dim rsSetup As New adodb.Recordset
Dim SQLQ
isTransferGP = False
If glbCompSerial = "S/N - 2259W" Then 'Ticket #19473
    SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & Product_Info & "' AND PARA_CATEGORY2='Integration Setup' AND PARA_NAME='" & xFunction & "' "
Else
    SQLQ = "SELECT * FROM APPLICATION_PARAMETER WHERE PARA_TYPE='Integration' AND PARA_CATEGORY='" & Product_Info & "' AND PARA_CATEGORY2='Integration_Setup2' AND PARA_NAME='" & xFunction & "' "
End If
rsSetup.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If rsSetup.EOF Then Exit Function
isTransferGP = rsSetup("PARA_VALUE") = "1"
rsSetup.Close
End Function

Public Sub ToRehireSub() 'Ticket #19937 for Samuel -  Franks 05/06/2011
        'If gSec_Inq_Rehire Then
        Screen.MousePointer = HOURGLASS
        'UnloadFrms
        Unload frmEEBASIC
        Screen.MousePointer = HOURGLASS
        Load frmEREHIRE
        'frmEREHIRE.ZOrder 0
        Screen.MousePointer = DEFAULT
        
        'Call UnloadFrms("newform")
End Sub

Public Function IsLinDupPayrollID(xEmpNo, xPayID, xNewHire, xIsAct, xSIN)
Dim rsEMP As New adodb.Recordset
Dim rsTmp As New adodb.Recordset
Dim SQLQ As String
Dim SQL2 As String
Dim xActiPayID As Long
Dim xTermPayID As Long
Dim xCurPayID As Long
Dim retval As Boolean
    retval = False

    SQLQ = "SELECT TOP 1 ED_EMPNBR, ED_SIN, ED_PAYROLL_ID FROM HREMP WHERE (1=1) "
    If xIsAct = "Y" Then
        If Not xNewHire = "Y" Then
            SQLQ = SQLQ & "AND NOT (ED_EMPNBR = " & xEmpNo & ") "
        End If
    End If
    SQLQ = SQLQ & "AND ED_PAYROLL_ID = '" & xPayID & "' "
    If rsEMP.State <> 0 Then rsEMP.Close
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        retval = True
    End If
    
    'not duplicate then check term table
    If retval = False Then '
        SQLQ = "SELECT TOP 1 ED_EMPNBR, ED_SIN, ED_PAYROLL_ID FROM Term_HREMP WHERE (1=1) "
        If xIsAct = "Y" Then
            'From Active
        Else
            'From Term
            If Len(xSIN) > 0 Then
                SQLQ = SQLQ & "AND NOT (ED_SIN = '" & xSIN & "') "
            End If
        End If
        SQLQ = SQLQ & "AND ED_PAYROLL_ID = '" & xPayID & "' "
        If rsEMP.State <> 0 Then rsEMP.Close
        rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsEMP.EOF Then
            retval = True
        End If
    End If
    If rsEMP.State <> 0 Then rsEMP.Close
    
    IsLinDupPayrollID = retval
End Function

Public Function getNextLinPayrollID(IsReset)  'Ticket #29759 Franks 02/15/2017
Dim rsPA As New adodb.Recordset
Dim rsEMP As New adodb.Recordset
Dim SQLQ As String
Dim xActiPayID As Long
Dim xTermPayID As Long
Dim xCurPayID As Long
Dim xNextPayIDCal As Long
Dim xNextPayIDTBL As Long

    xActiPayID = 0
    xTermPayID = 0
    
    'active table
    SQLQ = "SELECT TOP 1 ED_EMPNBR, ED_SIN, ED_PAYROLL_ID, CAST(ED_PAYROLL_ID AS INT) AS PayID_INT FROM HREMP WHERE NOT (ED_PAYROLL_ID IS NULL) "
    SQLQ = SQLQ & "AND ISNUMERIC(ED_PAYROLL_ID) = 1 " 'show records which it is numeric
    SQLQ = SQLQ & "ORDER BY PayID_INT DESC "
    If rsEMP.State <> 0 Then rsEMP.Close
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        xActiPayID = rsEMP("PayID_INT")
    End If
    
    'term table
    SQLQ = "SELECT TOP 1 ED_EMPNBR, ED_SIN, ED_PAYROLL_ID, CAST(ED_PAYROLL_ID AS INT) AS PayID_INT,TERM_SEQ, ED_ID FROM Term_HREMP WHERE NOT (ED_PAYROLL_ID IS NULL) "
    SQLQ = SQLQ & "AND ISNUMERIC(ED_PAYROLL_ID) = 1 " 'show records which it is numeric
    SQLQ = SQLQ & "ORDER BY PayID_INT DESC "
    If rsEMP.State <> 0 Then rsEMP.Close
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic
    If Not rsEMP.EOF Then
        xTermPayID = rsEMP("PayID_INT")
    End If
    
    If xActiPayID > xTermPayID Then
        xCurPayID = xActiPayID
    Else
        xCurPayID = xTermPayID
    End If
    
    xNextPayIDCal = xCurPayID + 1
    
    If IsReset = "Y" Then
        'get the latest next Payroll ID from the databse, and then update this field on the Company Master
        SQLQ = "UPDATE HRPARCO SET PC_NEXT_POS_NBR = " & xNextPayIDCal & " "
        gdbAdoIhr001.Execute SQLQ
        getNextLinPayrollID = xNextPayIDCal
    Else
        'get next Payroll ID
        rsPA.Open "SELECT * from HRPARCO", gdbAdoIhr001, adOpenStatic, adLockPessimistic
        xNextPayIDTBL = xNextPayIDCal
        If Not rsPA.EOF Then
            xNextPayIDTBL = rsPA("PC_NEXT_POS_NBR")
            If xNextPayIDTBL < xNextPayIDCal Then
                xNextPayIDTBL = xNextPayIDCal
            End If
            rsPA("PC_NEXT_POS_NBR") = xNextPayIDTBL + 1
            rsPA.Update
        End If
        rsPA.Close
        
        getNextLinPayrollID = xNextPayIDTBL
    End If
    
End Function
