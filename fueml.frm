VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmUEmergLeave 
   Appearance      =   0  'Flat
   Caption         =   "Mass Update Emergency Leave"
   ClientHeight    =   7710
   ClientLeft      =   1560
   ClientTop       =   2325
   ClientWidth     =   10410
   DrawMode        =   1  'Blackness
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7710
   ScaleWidth      =   10410
   Tag             =   "Emergency Leave Setup"
   WindowState     =   2  'Maximized
   Begin VB.TextBox Updstats 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      DataField       =   "EL_LDATE"
      Height          =   285
      Index           =   0
      Left            =   7200
      TabIndex        =   23
      Text            =   "LDATE"
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Updstats 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      DataField       =   "EL_LTIME"
      Height          =   285
      Index           =   1
      Left            =   8280
      TabIndex        =   22
      Text            =   "LTIME"
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Updstats 
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      DataField       =   "EL_LUSER"
      Height          =   285
      Index           =   2
      Left            =   9360
      TabIndex        =   21
      Text            =   "LUSER"
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin Threed.SSPanel panDetails 
      Height          =   2655
      Left            =   120
      TabIndex        =   12
      Top             =   3120
      Width           =   8535
      _Version        =   65536
      _ExtentX        =   15055
      _ExtentY        =   4683
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EL_EMP"
         Height          =   285
         Index           =   2
         Left            =   2130
         TabIndex        =   4
         Tag             =   "00-Enter Status Code"
         Top             =   1515
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDEM"
      End
      Begin INFOHR_Controls.CodeLookup clpPT 
         DataField       =   "EL_PT"
         Height          =   285
         Left            =   2130
         TabIndex        =   5
         Tag             =   "EDPT-Category"
         Top             =   1845
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDPT"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EL_ORG"
         Height          =   285
         Index           =   1
         Left            =   2130
         TabIndex        =   3
         Tag             =   "00-Enter Union Code"
         Top             =   1185
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDOR"
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "EL_LOC"
         Height          =   285
         Index           =   0
         Left            =   2130
         TabIndex        =   2
         Tag             =   "00-Enter Location Code"
         Top             =   855
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "EDLC"
      End
      Begin INFOHR_Controls.CodeLookup clpDept 
         DataField       =   "EL_DEPTNO"
         Height          =   285
         Left            =   2130
         TabIndex        =   1
         Tag             =   "00-Specific Department Desired"
         Top             =   525
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         MaxLength       =   7
         LookupType      =   2
      End
      Begin INFOHR_Controls.CodeLookup clpDiv 
         DataField       =   "EL_DIV"
         Height          =   285
         Left            =   2130
         TabIndex        =   0
         Tag             =   "00-Specific Division Desired"
         Top             =   195
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin MSMask.MaskEdBox medEml 
         DataField       =   "EL_EML"
         DataSource      =   "Data2"
         Height          =   285
         Left            =   2445
         TabIndex        =   6
         Tag             =   "20-Amount of entitlement during the period"
         Top             =   2175
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "0"
         PromptChar      =   "_"
      End
      Begin VB.Label lblLocation 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   20
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lblStatus 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   19
         Top             =   1560
         Width           =   450
      End
      Begin VB.Label lblUnion 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Union"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   18
         Top             =   1230
         Width           =   420
      End
      Begin VB.Label lblDept 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   570
         Width           =   825
      End
      Begin VB.Label lblDiv 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Division"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   16
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblSelectCrit 
         BackStyle       =   0  'Transparent
         Caption         =   "Selection Criteria"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lblPT 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   1890
         Width           =   630
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Emergency Leave (Days)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   2220
         Width           =   2130
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   615
      Left            =   0
      TabIndex        =   11
      Top             =   7095
      Width           =   10410
      _Version        =   65536
      _ExtentX        =   18362
      _ExtentY        =   1085
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   0
      BorderWidth     =   0
      BevelOuter      =   0
      Font3D          =   1
      Alignment       =   1
      Begin VB.CommandButton cmdReCalcAll 
         Caption         =   "&ReCalculate All"
         Height          =   375
         Left            =   5640
         TabIndex        =   10
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton cmdReCalc 
         Caption         =   "&ReCalculate"
         Height          =   375
         Left            =   4080
         TabIndex        =   9
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton cmdUpdateAll 
         Caption         =   "Update All"
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   120
         Width           =   1500
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1500
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8640
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fueml.frx":0000
      Height          =   2625
      Left            =   0
      OleObjectBlob   =   "fueml.frx":0014
      TabIndex        =   24
      Tag             =   "Overtime Master"
      Top             =   120
      Width           =   10410
   End
End
Attribute VB_Name = "frmUEmergLeave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim EECri As String, OneSet%, X%
Dim XUpdCount, Actn
Dim SQLQ, strFld
Dim SaveHours
Dim rsDATA As New ADODB.Recordset
Dim fglbNew As Boolean

Private Sub chkCOEFlag_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Function chkFUEML()

Dim SQLQ As String, msg$, dd&, Response%, X%
Dim DgDef As Variant, Title$, DCurPDate As Variant

chkFUEML = False

On Error GoTo chkFUEML_Err
If Not IsNumeric(medEml.Text) Then
    MsgBox "Emergency Leave Must be numeric"
    medEml.Text = ""
    medEml.SetFocus
    Exit Function
End If

If Len(medEml.Text) < 1 Then
    MsgBox "Emergency Leave is a required field"
    medEml.SetFocus
    Exit Function
End If


If Len(clpDiv.Text) > 0 And clpDiv.Caption = "Unassigned" Then
    MsgBox lStr("If Division Entered - it must be known")
     clpDiv.SetFocus
    Exit Function
End If

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox "If Department Entered - it must be known"
     clpDept.SetFocus
    Exit Function
End If

For X% = 0 To 2
    If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
        MsgBox "If code entered it must be known"
        clpCode(X%).SetFocus
        Exit Function
    End If
Next X%

If Len(clpPT.Text) > 0 And clpPT.Caption = "Unassigned" Then
    MsgBox lStr("Category code must be valid")
     clpPT.SetFocus
    Exit Function
End If

chkFUEML = True

Exit Function

chkFUEML_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkFUEML", "HR_EMLSETUP", "CHKFUEML")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

Public Sub cmdClose_Click()
Unload Me
End Sub

Public Sub cmdDelete_Click()
    Dim A As Integer, msg As String, X%


    If Data1.Recordset.BOF And Data1.Recordset.EOF Then
      MsgBox "Nothing to Delete"
      Exit Sub
    End If
    On Error GoTo Del_Err
    
    msg = "Are You Sure You Want To Delete "
    msg = msg & "This Record?"
    A% = MsgBox(msg, vbQuestion + vbYesNo, "Confirm Delete")
    If A% = vbNo Then Exit Sub
    
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete adAffectCurrent
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
    Call Display_Value
    
    fglbNew = False
    
    Exit Sub
Del_Err:
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_EMLSETUP", "Delete")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        Resume Next
    Else
        Unload Me
    End If

End Sub

Public Sub cmdNew_Click()
Dim SQLQ As String, msg$, X%
Dim Title$, DgDef As Variant, Response%
On Error GoTo AddN_Err
fglbNew = True
Call SET_UP_MODE
If Not gSec_Upd_Other_Entitlements Then
    MsgBox "You Do Not Have Authority For This Transaction"
    Exit Sub
End If


Call Set_Control("B", Me)

Screen.MousePointer = DEFAULT

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "ATTEND", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub

Private Sub cmdReCalc_Click()
    'Ticket #22682 - Release 8.0
    Call CalcEMLTaken
End Sub

Private Sub cmdReCalcAll_Click()

    'Ticket #22682 - Release 8.0
    If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        Data1.Recordset.MoveFirst
        Do
            Call Display_Value
            Call CalcEMLTaken
            
            Data1.Recordset.MoveNext
        Loop Until Data1.Recordset.EOF
    End If
    
    Data1.Refresh
    
    Call Display_Value
    
    Screen.MousePointer = DEFAULT
        
    Exit Sub

End Sub

Private Sub cmdUpdate_Click()
    Dim msg$, DgDef As Variant, Response%
    Dim dd&
    Dim Title$
    Dim recCount As Integer
    
    If Not gSec_Upd_Other_Entitlements Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If
    
    Actn = "A"
    If Not chkFUEML() Then Exit Sub
    
    Title$ = "Emergency Leave Mass Update"
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
    msg$ = "Are you sure you want to Update Emergency Leave for this criteria?"
    Response% = MsgBox(msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If
    
    recCount = getRecordCount_Add
    If recCount > 0 Then
        msg$ = Str(recCount)
        If recCount = 1 Then msg$ = msg$ & " employee's Emergency Leave " Else msg$ = msg$ & " employee's Emergency Leave "
        msg$ = msg$ & "will be Updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
        Response% = MsgBox(msg$, DgDef, Title)    ' Get user response.
        If Response = IDNO Then
            Exit Sub
        End If
    Else
        MsgBox "No Employee record found to update with Emergency Leave."
        Exit Sub
    End If
    
    Screen.MousePointer = HOURGLASS
    
    If Not modInsRecs() Then
        Screen.MousePointer = DEFAULT
        Exit Sub
    End If
    
    Screen.MousePointer = DEFAULT
    If XUpdCount > 0 Then
        MsgBox Str(XUpdCount) & " Records Updated Successfully"
    Else
        MsgBox "No Records Added"
    End If
End Sub


Private Sub cmdUpdateAll_Click()
Dim msg$, DgDef As Variant, Response%
Dim dd&
Dim Title$
Dim recCount As Integer
Dim failed As String
Dim c As Long

On Error GoTo Mod_Err
    
    'Ticket #22682 - Release 8.0
    If Not gSec_Upd_Other_Entitlements Then
        MsgBox "You Do Not Have Authority For This Transaction"
        Exit Sub
    End If
    
    Actn = "A"
    
    Title$ = "Emergency Leave Mass Update"
    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2   ' Describe dialog.
    msg$ = "Are you sure you want to Update Emergency Leave for this criteria?"
    Response% = MsgBox(msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then    ' Evaluate response
        Exit Sub
    End If

    failed = ""
    c = 1

    If Data1.Recordset.EOF = False And Data1.Recordset.BOF = False Then
        Data1.Recordset.MoveFirst
        Do
            Call Display_Value
            
            If chkFUEML() Then
                'Get # of records to update and display with the prompt
                recCount = getRecordCount_Add
                If recCount > 0 Then
                    msg$ = Str(recCount)
                    If recCount = 1 Then msg$ = msg$ & " employee's Emergency Leave " Else msg$ = msg$ & " employee's Emergency Leave "
                    msg$ = msg$ & "will be updated. " & vbCrLf & vbCrLf & "Do you want to proceed?"
                    Response% = MsgBox(msg$, DgDef, Title)    ' Get user response.
                    If Response = IDNO Then
                        c = c - 1
                        GoTo nextRule
                    End If
                Else
                    MsgBox "Employees for this selection do not exist!"
                    
                    GoTo nextRule
                End If
                
                Screen.MousePointer = HOURGLASS
                
                'Update Employee's Emergency Leave
                If Not modInsRecs() Then
                    'Failed
                    failed = failed & "Rule " & CStr(c) & ": "
                    If Not IsNull(Data1.Recordset("EL_DIV")) Then failed = failed & Data1.Recordset("EL_DIV") & ", "
                    If Not IsNull(Data1.Recordset("EL_DEPTNO")) Then failed = failed & Data1.Recordset("EL_DEPTNO") & ", "
                    If Not IsNull(Data1.Recordset("EL_LOC")) Then failed = failed & Data1.Recordset("EH_LOC") & ", "
                    If Not IsNull(Data1.Recordset("EL_ORG")) Then failed = failed & Data1.Recordset("EL_ORG") & ", "
                    If Not IsNull(Data1.Recordset("EL_EMP")) Then failed = failed & Data1.Recordset("EL_EMP") & ", "
                    If Not IsNull(Data1.Recordset("EL_PT")) Then failed = failed & Data1.Recordset("EL_PT") & ", "
                    If Not IsNull(Data1.Recordset("EL_EML")) Then failed = failed & Data1.Recordset("EL_EML") & ", "
                    failed = Left(failed, Len(failed) - 2) & vbCrLf
                End If

                Screen.MousePointer = DEFAULT
            Else
                'Failed
                failed = failed & "Rule " & CStr(c) & ": "
                If Not IsNull(Data1.Recordset("EL_DIV")) Then failed = failed & Data1.Recordset("EL_DIV") & ", "
                If Not IsNull(Data1.Recordset("EL_DEPTNO")) Then failed = failed & Data1.Recordset("EL_DEPTNO") & ", "
                If Not IsNull(Data1.Recordset("EL_LOC")) Then failed = failed & Data1.Recordset("EH_LOC") & ", "
                If Not IsNull(Data1.Recordset("EL_ORG")) Then failed = failed & Data1.Recordset("EL_ORG") & ", "
                If Not IsNull(Data1.Recordset("EL_EMP")) Then failed = failed & Data1.Recordset("EL_EMP") & ", "
                If Not IsNull(Data1.Recordset("EL_PT")) Then failed = failed & Data1.Recordset("EL_PT") & ", "
                If Not IsNull(Data1.Recordset("EL_EML")) Then failed = failed & Data1.Recordset("EL_EML") & ", "
                failed = Left(failed, Len(failed) - 2) & vbCrLf
            End If
nextRule:
            c = c + 1
            Data1.Recordset.MoveNext
        Loop Until Data1.Recordset.EOF
    End If
    
    Data1.Refresh
    
    Call Display_Value
    
    Screen.MousePointer = DEFAULT
    
    If Len(failed) = 0 And c > 1 Then
        MsgBox "All Rules applied. Emergency Leave mass update completed successfully.", vbInformation + vbOKOnly, "Emergency Leave Mass Update"
    ElseIf Len(failed) = 0 And c = 1 Then
        MsgBox "No Rules applied.", vbInformation + vbOKOnly, "Emergency Leave Mass Update"
    Else
        MsgBox "The Emergency Leave update for the following Rules failed:" & vbCrLf & failed, vbInformation + vbOKOnly, "Emergency Leave Mass Update"
    End If
    
    Exit Sub
    
Mod_Err:
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdateAll", "Mass Update Emergency Leave", "Modify")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
         RollBack
        Resume Next
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMUEMERGLEAVE"
End Sub

Private Sub Form_Load()
glbOnTop = "FRMUEMERGLEAVE"
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "SELECT ID,EL_DIV,EL_DEPTNO,EL_LOC,EL_ORG,EL_EMP,EL_PT,EL_EML,EL_LDATE,EL_LTIME,EL_LUSER FROM HR_EMLSETUP"
Data1.Refresh

Call setCaption(lblDiv)
Call setCaption(lblDept)
Call setCaption(lblLocation)
Call setCaption(lblUnion)
Call setCaption(lblStatus)
Call setCaption(lblPT)

Call setCaption(vbxTrueGrid.Columns(0))
Call setCaption(vbxTrueGrid.Columns(1))
Call setCaption(vbxTrueGrid.Columns(2))
Call setCaption(vbxTrueGrid.Columns(3))
Call setCaption(vbxTrueGrid.Columns(4))
Call setCaption(vbxTrueGrid.Columns(5))

Screen.MousePointer = HOURGLASS
Call setRptCaption(Me)
'If glbCompSerial = "S/N - 2227W" Then clpCode(4).MaxLength = 6

Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmUEmergLeave = Nothing
End Sub

Private Function modInsRecs()

Dim msg$, DgDef As Variant, Response%, noRecs&
Dim rsTA As New ADODB.Recordset

modInsRecs = False
On Error GoTo cmdInsErr

    XUpdCount = 0
    SQLQ = "SELECT COUNT(ED_EMPNBR) as empcnt FROM HREMP " & WSQLQ1
    rsTA.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    
    If rsTA.EOF = False And rsTA.BOF = False Then
        XUpdCount = rsTA("empcnt")
    
    
        SQLQ = "UPDATE HREMP SET"
        SQLQ = SQLQ & " ED_EML=" & CDbl(medEml)
        SQLQ = SQLQ & WSQLQ1
        gdbAdoIhr001.Execute SQLQ
    End If
    rsTA.Close
    Call CalcEMLTaken
    
modInsRecs = True

Exit Function
cmdInsErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
If glbErrNum& = -2147467259 Then
    MsgBox "The changes were not successful because it would create duplicate values."
    Exit Function
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EML Error", "HR_EMLSETUP", "Insert")
    Screen.MousePointer = DEFAULT
    If gintRollBack% = False Then
        RollBack
        Resume Next
    Else
        Unload Me
    End If
End If
End Function

Private Sub CalcEMLTaken()
    Dim rsATT As New ADODB.Recordset
    'Dim rsEMP As New ADODB.Recordset
    Dim SQLQ
    Dim toteml
    Dim xRows As Long
    Dim xRow As Long
        
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = HOURGLASS
    
    gdbAdoIhr001.BeginTrans
        
        If glbtermopen Then
            SQLQ = "SELECT Sum(AD_HRS) AS SumOfAD_HRS, ED_EMPNBR, Sum(HREMP.ED_DHRS) AS SumOfED_DHRS"
            SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
            SQLQ = SQLQ & " GROUP BY Year((AD_DOA)),AD_EMELEA,ED_EMPNBR"
            If glbOracle Then
                SQLQ = SQLQ & " HAVING TO_CHAR(AD_DOA,'YYYY')  = " & Year(Date)
            Else
                SQLQ = SQLQ & " HAVING YEAR(AD_DOA) = " & Year(Date)
            End If
            SQLQ = SQLQ & " AND ((AD_EMELEA)<>0)"
            rsATT.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            SQLQ = "SELECT Sum(AD_HRS) AS SumOfAD_HRS, ED_EMPNBR, Sum(HREMP.ED_DHRS) AS SumOfED_DHRS"
            SQLQ = SQLQ & " FROM HREMP INNER JOIN HR_ATTENDANCE ON HREMP.ED_EMPNBR = HR_ATTENDANCE.AD_EMPNBR"
            If glbOracle Then
                SQLQ = SQLQ & " GROUP BY TO_CHAR(AD_DOA,'YYYY') ,AD_EMELEA,ED_EMPNBR"
                SQLQ = SQLQ & " HAVING TO_CHAR(AD_DOA,'YYYY')  = " & Year(Date)
            Else
                SQLQ = SQLQ & " GROUP BY Year((AD_DOA)),AD_EMELEA,ED_EMPNBR"
                SQLQ = SQLQ & " HAVING YEAR(AD_DOA) = " & Year(Date)
            End If
            SQLQ = SQLQ & " AND ((AD_EMELEA)<>0)"
            
            rsATT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
        
        If Not rsATT.EOF Then
            MDIMain.panHelp(0).FloodType = 1
            MDIMain.panHelp(1).Caption = " Please Wait" '
            MDIMain.panHelp(2).Caption = ""             '
            'MDIMain.panHelp(0).FloodPercent = 10
            
            'UPDATE THE HREMP WITH EML TAKEN
            rsATT.MoveFirst
            xRows = rsATT.RecordCount
            xRow = 0
            Do While Not rsATT.EOF
                MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
                
                toteml = 0
                toteml = toteml + CDbl(rsATT("SumOfAD_HRS"))
               
                If glbOracle Then
                    SQLQ = "Update HREMP"
                    SQLQ = SQLQ & " SET HREMP.ED_EMLT =" & toteml
                    SQLQ = SQLQ & WSQLQ1
                    SQLQ = SQLQ & " AND ED_EMPNBR=" & rsATT("ED_EMPNBR")
                    gdbAdoIhr001.Execute SQLQ
                Else
                    SQLQ = "Update HREMP"
                    SQLQ = SQLQ & " SET ED_EMLT = " & toteml
                    SQLQ = SQLQ & WSQLQ1
                    SQLQ = SQLQ & " AND ED_EMPNBR=" & rsATT("ED_EMPNBR")
                    gdbAdoIhr001.Execute SQLQ
                End If
               
                xRow = xRow + 1
                rsATT.MoveNext
            Loop
        End If
        rsATT.Close
        'rsEMP.Close
            
    'MDIMain.panHelp(0).FloodPercent = 30
    
    gdbAdoIhr001.CommitTrans
    
    Screen.MousePointer = DEFAULT
    
    MDIMain.panHelp(0).FloodPercent = 100
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = ""
    MDIMain.panHelp(2).Caption = ""
    
    Exit Sub

ErrorHandler:
Screen.MousePointer = DEFAULT
glbFrmCaption$ = "EML Entitlement Recalculation"
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "CalcEMLTaken", "", "EMLTAKEN")
If gintRollBack% = False Then
    Resume Next
End If
End Sub

Private Function WSQLQ1() As String
'COMMENTED BY SAM AS IT DOES NOT MATCH THE COLUMNS IN EML TABLE
'ucommented by bryan, where does this criteria affect EML table??

WSQLQ1 = WSQLQ1 & " WHERE " & glbSeleDeptUn
If Len(clpDept.Text) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_DEPTNO = '" & clpDept.Text & "'"
If Len(clpDiv.Text) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_DIV = '" & clpDiv.Text & "' "
If Len(clpCode(0).Text) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_LOC = '" & clpCode(0).Text & "' "
If Len(clpCode(1).Text) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_ORG = '" & clpCode(1).Text & "' "
If Len(clpCode(2).Text) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_EMP = '" & clpCode(2).Text & "' "
If Len(clpPT.Text) > 0 Then WSQLQ1 = WSQLQ1 & " AND ED_PT = '" & clpPT.Text & "' "

End Function

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf Data1.Recordset.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
clpDiv.Enabled = TF
clpDept.Enabled = TF
clpPT.Enabled = TF
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
medEml.Enabled = TF
If Not Data1.Recordset.EOF Then
    cmdUpdate.Enabled = True
Else
    cmdUpdate.Enabled = False
End If

End Sub

Public Property Get RelateMode() As RelateModeEnum
    RelateMode = nothingrelate
End Property

Public Property Get UpdateRight() As Boolean
    UpdateRight = gSec_Upd_Other_Entitlements
End Property

Public Property Get Addable() As Boolean
    Addable = True
End Property

Public Property Get Updateble() As Boolean
    Updateble = True
End Property

Public Property Get Deleteble() As Boolean
    Deleteble = True
End Property

Public Property Get Printable() As Boolean
    Printable = False
End Property

Public Property Get ChangeAction() As UpdateStateEnum
If fglbNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property

Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fglbNew = True
End Property

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
       
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT ID,EL_DIV,EL_DEPTNO,EL_LOC,EL_ORG,EL_EMP,EL_PT,EL_EML,EL_LDATE,EL_LTIME,EL_LUSER FROM HR_EMLSETUP"

    Data1.RecordSource = SQLQ
    Data1.Refresh

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Call Display_Value
End Sub

Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        SQLQ = "SELECT ID,EL_DIV,EL_DEPTNO,EL_LOC,EL_ORG,EL_EMP,EL_PT,EL_EML,EL_LDATE,EL_LTIME,EL_LUSER FROM HR_EMLSETUP"
        SQLQ = SQLQ & " WHERE ID = " & Data1.Recordset!ID
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
        Call Set_Control("R", Me, rsDATA)
    
    End If
    Call SET_UP_MODE
End Sub

Sub cmdOK_Click()
Dim xID As Long

On Error GoTo Add_Err

    
If chkFUEML = False Then Exit Sub

rsDATA.Requery

If fglbNew Then rsDATA.AddNew

Call UpdUStats(Me)
        
gdbAdoIhr001.BeginTrans
Call Set_Control("U", Me, rsDATA)
rsDATA.Update
gdbAdoIhr001.CommitTrans
rsDATA.Resync
xID = rsDATA!ID
Data1.Refresh


Data1.Recordset.Find "ID=" & xID

fglbNew = False

Call SET_UP_MODE

Exit Sub

Add_Err:
If Err = 3022 Then
    Data1.Recordset.CancelUpdate    ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRBENFT", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Sub cmdCancel_Click()
Dim X, bk
On Error GoTo Can_Err


Call Display_Value
fglbNew = False
Call SET_UP_MODE

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_BENFTS_GROUP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Function getRecordCount_Add()
    Dim SQLQ As String
    Dim rsEMP As New ADODB.Recordset
    Dim recCount As Integer
    
    getRecordCount_Add = 0
    recCount = 0

    SQLQ = "SELECT COUNT(ED_EMPNBR) as TOT_REC FROM HREMP " & WSQLQ1
    rsEMP.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic, adCmdText
    If Not rsEMP.EOF Then
        recCount = rsEMP("TOT_REC")
    Else
        recCount = 0
    End If
    rsEMP.Close
    Set rsEMP = Nothing
    
    getRecordCount_Add = recCount

End Function

