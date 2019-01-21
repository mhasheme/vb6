VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEWorkFlow 
   Caption         =   "Work Flow Overview"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   13020
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdA1 
      Caption         =   "Task"
      Height          =   375
      Index           =   6
      Left            =   9360
      TabIndex        =   10
      Top             =   120
      Width           =   3495
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Target (Days)"
      Height          =   375
      Index           =   5
      Left            =   8160
      TabIndex        =   9
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Step"
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstALL 
      Height          =   4335
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   480
      Width           =   12945
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Cmptd"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Employee"
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Plant"
      Height          =   375
      Index           =   2
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdA1 
      Caption         =   "Type"
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   6705
      Width           =   13020
      _Version        =   65536
      _ExtentX        =   22966
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      Begin VB.CommandButton cmdClearAll 
         Appearance      =   0  'Flat
         Caption         =   "&Clear All Cmptd."
         Height          =   375
         Left            =   1920
         TabIndex        =   12
         Tag             =   "Mark all messages as being completed."
         Top             =   0
         Width           =   1620
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete Cmptd Work Flow"
         Height          =   375
         Left            =   3600
         TabIndex        =   11
         Tag             =   "Mark all messages as being completed."
         Top             =   0
         Width           =   2100
      End
      Begin VB.CommandButton cmdMarkAll 
         Appearance      =   0  'Flat
         Caption         =   "&Mark All Cmptd."
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Tag             =   "Mark all messages as being completed."
         Top             =   0
         Width           =   1620
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         Height          =   375
         Left            =   5760
         TabIndex        =   6
         Tag             =   "Save the changes made"
         Top             =   0
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   9120
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
End
Attribute VB_Name = "frmEWorkFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xPlanOrderBy As String, xEDateOrderBy As String, xEmpOrderBy As String, xTypeOrderBy As String

Private Sub cmdA1_Click(Index As Integer)
    If Index = 1 Then 'Name
        If xEmpOrderBy = "ASC" Then
            xEmpOrderBy = "DESC"
        Else
            xEmpOrderBy = "ASC"
        End If
        Call RetrieveData(1)
    End If
    If Index = 2 Then 'Plant
        If xPlanOrderBy = "ASC" Then
            xPlanOrderBy = "DESC"
        Else
            xPlanOrderBy = "ASC"
        End If
        Call RetrieveData(2)
    End If
    If Index = 3 Then 'Type
        If xTypeOrderBy = "ASC" Then
            xTypeOrderBy = "DESC"
        Else
            xTypeOrderBy = "ASC"
        End If
        Call RetrieveData(3)
    End If

End Sub

Private Sub cmdClearAll_Click()
Dim x As Integer
For x = 0 To lstALL.ListCount - 1
    lstALL.selected(x) = False
Next
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim SQLQ As String, Msg$, Response%, Title$, DgDef As Variant
Dim xName As String, xSurName, xFName, xEDate, xSec, xSecDesc As String, xType As String, xStep
Dim xTemp As String
Dim rsWorkFlow As New ADODB.Recordset
Dim x As Integer
Dim I As Integer

On Error GoTo MAll_Err

I = 0
For x = 0 To lstALL.ListCount - 1
    If lstALL.selected(x) = True Then
        I = I + 1
    End If
Next
If I = 0 Then
    MsgBox "No Cmptd Selected"
    Exit Sub
End If

Msg$ = lStr("Delete all selected Work Flow records?")
'Msg = Msg$ & Chr(10) & "listed as completed?"
Title$ = "Delete all completed?"   ' zzz
DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2  ' Describe dialog.
Response% = MsgBox(Msg, DgDef, Title)    ' Get user response.
If Response = IDYES Then    ' Evaluate response
    Screen.MousePointer = HOURGLASS
    
    For x = 0 To lstALL.ListCount - 1
        If lstALL.selected(x) Then
            'MsgBox Left(lstALL.List(X), 100)
            xSecDesc = Trim(Mid(lstALL.List(x), 9, 33))
            xSec = getCodeFromDesc("EDSE", xSecDesc)
            
            xName = Mid(lstALL.List(x), 42, 43)
            xSurName = Trim(CSVGet(xName, 1))
            xFName = Trim(CSVGet(xName, 2))
            
            xType = Trim(Mid(lstALL.List(x), 85, 45))
            xType = getCodeFromDesc("WKFL", xType)
            xStep = Trim(Mid(lstALL.List(x), 130, 10))
            
            
            
            SQLQ = "SELECT * FROM HRWORKFLOW_EMPLOYEE WHERE (1=1) "
            SQLQ = SQLQ & "AND PE_SURNAME = '" & xSurName & "' "
            SQLQ = SQLQ & "AND PE_FNAME = '" & xFName & "' "
            SQLQ = SQLQ & "AND PE_SECTION = '" & xSec & "' "
            SQLQ = SQLQ & "AND PE_STEP = " & xStep & " "
            SQLQ = SQLQ & "AND PE_WORKFLOW = '" & xType & "' "
            If rsWorkFlow.State <> 0 Then rsWorkFlow.Close
            rsWorkFlow.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsWorkFlow.EOF Then
                rsWorkFlow("PE_CMPTD_FLAG") = 1
                rsWorkFlow("PE_CMPTD_DATE") = Date
                rsWorkFlow.Update
            End If
            rsWorkFlow.Close
            
        End If
    Next
    Call RetrieveData
    Screen.MousePointer = DEFAULT
    
End If


Exit Sub

MAll_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdMarkAll", "HRWORKFLOW_EMPLOYEE", "Mark All")

End Sub

Private Sub cmdMarkAll_Click()
Dim x As Integer
For x = 0 To lstALL.ListCount - 1
    lstALL.selected(x) = True
Next

End Sub



Private Sub Form_Activate()
glbOnTop = "frmEWorkFlow"
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
    glbOnTop = "frmEWorkFlow"
    xPlanOrderBy = "ASC"
    xEDateOrderBy = "ASC"
    xEmpOrderBy = "ASC"
    xTypeOrderBy = "ASC"
    Call RetrieveData
End Sub
Private Sub RetrieveData(Optional xInx)
Dim rsWorkFlow As New ADODB.Recordset
Dim SQLQ As String
Dim xStr As String
Dim xTemp As String
    lstALL.Clear
    SQLQ = "SELECT * FROM HRWORKFLOW_EMPLOYEE WHERE (1=1) "
    SQLQ = SQLQ & "AND " & glbSelePESection & " "
    SQLQ = SQLQ & "AND PE_CMPTD_FLAG = 0 "
    If IsMissing(xInx) Then
        SQLQ = SQLQ & "ORDER BY PE_SURNAME,PE_FNAME "
    Else
        If xInx = 1 Then
            SQLQ = SQLQ & "ORDER BY PE_SURNAME + PE_FNAME " & xEmpOrderBy
        End If
        If xInx = 2 Then
            SQLQ = SQLQ & "ORDER BY PE_SECTION " & xPlanOrderBy & ", PE_SURNAME + PE_FNAME"
        End If
        If xInx = 3 Then
            SQLQ = SQLQ & "ORDER BY PE_WORKFLOW " & xTypeOrderBy & ", PE_SURNAME + PE_FNAME"
        End If
    End If

    rsWorkFlow.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    Do Until rsWorkFlow.EOF
        'xStr = Space(8) & Left(Trim(rsWorkFlow("PE_SURNAME")) & ", " & Trim(rsWorkFlow("PE_FNAME")) & Space(43), 43) 'Name
        xStr = Space(8)
        xTemp = GetTABLDesc("EDSE", rsWorkFlow("PE_SECTION"))
        xTemp = Trim(xTemp)
        xStr = xStr & Left(xTemp & Space(33), 33) 'Plant
        'xTemp = CVDate(rsWorkFlow("PE_EVENT_DATE"))
        'xStr = xStr & Left(xTemp & Space(20), 20) 'Event Date
        xTemp = Left(Trim(rsWorkFlow("PE_SURNAME")) & ", " & Trim(rsWorkFlow("PE_FNAME")) & Space(43), 43) 'Name
        xStr = xStr & Left(xTemp & Space(43), 43)
        xTemp = GetTABLDesc("WKFL", rsWorkFlow("PE_WORKFLOW"))
        xStr = xStr & Left(xTemp & Space(45), 45) 'Event Type
        xTemp = Trim(Str(rsWorkFlow("PE_STEP")))
        xStr = xStr & Left(xTemp & Space(25), 25) 'step
        xTemp = ""
        If Not IsNull(rsWorkFlow("PE_TARGET")) Then
            xTemp = Trim(Str(rsWorkFlow("PE_TARGET")))
        End If
        xStr = xStr & Left(xTemp & Space(20), 20) 'target
        
        xTemp = rsWorkFlow("PE_TASK")
        xStr = xStr & Left(xTemp & Space(100), 100)
        
        lstALL.AddItem xStr
        rsWorkFlow.MoveNext
    Loop
    rsWorkFlow.Close
    'lstALL.SetFocus
    
End Sub

Public Property Get ChangeAction() As UpdateStateEnum
If fGLBNew Then
    ChangeAction = NewRecord
Else
    ChangeAction = OPENING
End If
End Property
Public Property Let ChangeAction(vData As UpdateStateEnum)
If vData = NewRecord Then fGLBNew = True
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP 'MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = False  'True 'False
End Property

Public Property Get Addable() As Boolean
Addable = False  'True
End Property
Public Property Get Updateble() As Boolean
Updateble = False  'True
End Property
Public Property Get Deleteble() As Boolean
Deleteble = False  'True
End Property

Public Property Get Printable() As Boolean
Printable = False  'True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
'If fglbNew Then
'    UpdateState = NewRecord
'    TF = True
'ElseIf rsData.EOF Then
'    UpdateState = NoRecord
'    TF = False
'Else
    UpdateState = OPENING
    TF = True
'End If

Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
If Not Updateble Then TF = False

'fraEmpLookup.Visible = fglbNew
'txtEmpnbr.Visible = Not fglbNew

End Sub

