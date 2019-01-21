VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmEESTATSComm 
   Caption         =   "EMP Status/Dates Comments"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      DataSource      =   " "
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Tag             =   "00-Comments - free form"
      Top             =   840
      Width           =   6855
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   3885
      Width           =   7125
      _Version        =   65536
      _ExtentX        =   12568
      _ExtentY        =   952
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
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Tag             =   "Close and exit this screen"
         Top             =   120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Tag             =   "Close and exit this screen"
         Top             =   120
         Width           =   975
      End
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   525
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7125
      _Version        =   65536
      _ExtentX        =   12568
      _ExtentY        =   926
      _StockProps     =   15
      ForeColor       =   255
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
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1200
         TabIndex        =   7
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "lblEEName"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2880
         TabIndex        =   6
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee#"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   945
      End
   End
   Begin VB.Label lblTitle 
      Caption         =   "Comments"
      Height          =   255
      Index           =   19
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmEESTATSComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    If glbLOAComments = False Then
        frmEESTATS.txtEmpComm.Text = txtComments.Text
    End If
    
Unload Me

End Sub

Private Sub cmdUpdate_Click()
    Dim rsHRStatus As New ADODB.Recordset
    Dim SQLQ As String
    
    'glbLEE_ID, frmEESTATS.clpCode(1).Text, frmEESTATS.dlpDate(15).Text, frmEESTATS.dlpDate(16).Text
    SQLQ = "SELECT * FROM HRSTATUS "
    If IsDate(frmEESTATS.dlpDate(15).Text) Or IsDate(frmEESTATS.dlpDate(16).Text) Then
        SQLQ = SQLQ & " WHERE SC_REASON IN ('LOA') AND SC_EMPNBR=" & glbLEE_ID
        SQLQ = SQLQ & " AND SC_FDATE=" & Date_SQL(frmEESTATS.dlpDate(15).Text)
        SQLQ = SQLQ & " AND SC_TDATE=" & Date_SQL(frmEESTATS.dlpDate(16).Text)
        SQLQ = SQLQ & " AND SC_NEWEMP='" & frmEESTATS.clpCode(1).Text & "'"
        rsHRStatus.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRStatus.EOF Then
            rsHRStatus("SC_COMMENT") = txtComments.Text
            rsHRStatus.Update
        End If
        rsHRStatus.Close
        Set rsHRStatus = Nothing
    End If

End Sub

Private Sub Form_Activate()
glbOnTop = "frmEESTATSComm"
End Sub

Private Sub Form_GotFocus()
glbOnTop = "frmEESTATSComm"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim rsTERM As New ADODB.Recordset
Dim x%, SQLQ

glbOnTop = "frmEIncidentDemo"


Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = frmEESTATS.lblEENum

If glbLOAComments = False Then
    If EERetrieve() = False Then Exit Sub
Else
    'Release 8.1
    'Retrieve LOA Comments
    txtComments.Text = Get_LOA_Comments(glbLEE_ID, frmEESTATS.clpCode(1).Text, frmEESTATS.dlpDate(15).Text, frmEESTATS.dlpDate(16).Text)
    cmdUpdate.Visible = True
End If

MDIMain.panHelp(1).Caption = " "

End Sub

Private Function EERetrieve()
Dim rsORG As New ADODB.Recordset
Dim SQLQ As String
EERetrieve = False

On Error GoTo EERError

Screen.MousePointer = HOURGLASS


txtComments.Text = frmEESTATS.txtEmpComm.Text

EERetrieve = True

Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EMP Status/Dates Comments", "HREMP_OTHER", "SELECT")

Resume Next

Exit Function
'
End Function

Private Function Get_LOA_Comments(xEmpnbr, xReason, xFromDate, xToDate) As String
    Dim rsHRStatus As New ADODB.Recordset
    Dim SQLQ As String
    
    Screen.MousePointer = DEFAULT
    
    Get_LOA_Comments = ""
    
    SQLQ = "SELECT * FROM HRSTATUS "
    If IsDate(xFromDate) Or IsDate(xToDate) Then
        SQLQ = SQLQ & " WHERE SC_REASON IN ('LOA') AND SC_EMPNBR=" & xEmpnbr
        SQLQ = SQLQ & " AND SC_FDATE=" & Date_SQL(xFromDate)
        SQLQ = SQLQ & " AND SC_TDATE=" & Date_SQL(xToDate)
        SQLQ = SQLQ & " AND SC_NEWEMP='" & xReason & "'"
        rsHRStatus.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsHRStatus.EOF Then
            Get_LOA_Comments = IIf(IsNull(rsHRStatus("SC_COMMENT")), "", rsHRStatus("SC_COMMENT"))
        Else
            Get_LOA_Comments = ""
        End If
        rsHRStatus.Close
        Set rsHRStatus = Nothing
    Else
        Get_LOA_Comments = ""
    End If
    
End Function

