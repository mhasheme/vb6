VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmFollowUpList 
   Caption         =   "Follow Up List"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5145
   ScaleWidth      =   9690
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAllCompleted 
      Appearance      =   0  'Flat
      Caption         =   "Mark All Completed"
      Height          =   420
      Left            =   7320
      TabIndex        =   9
      Tag             =   "Save changes made"
      Top             =   3840
      Width           =   1875
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9690
      _Version        =   65536
      _ExtentX        =   17092
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblEEProdLine 
         AutoSize        =   -1  'True
         Caption         =   "Product Line"
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
         Left            =   6360
         TabIndex        =   8
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEEID 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "lblEEID"
         DataSource      =   "Data1"
         ForeColor       =   &H008080FF&
         Height          =   180
         Left            =   4680
         TabIndex        =   7
         Top             =   120
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblEEName 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         TabIndex        =   5
         Top             =   135
         Width           =   720
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
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
         TabIndex        =   4
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEENumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   160
         Width           =   1005
      End
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   4590
      Width           =   9690
      _Version        =   65536
      _ExtentX        =   17092
      _ExtentY        =   979
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
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
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
         Left            =   2340
         TabIndex        =   1
         Tag             =   "Cancel changes made"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&Save Changes"
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
         TabIndex        =   0
         Tag             =   "Save changes made"
         Top             =   0
         Width           =   1995
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   4080
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
         Caption         =   "Ado1"
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
      Bindings        =   "ffollowuplist.frx":0000
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "ffollowuplist.frx":0014
      TabIndex        =   10
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click the Row to mark/unmark as Completed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   3840
      Width           =   4455
   End
End
Attribute VB_Name = "frmFollowUpList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAllCompleted_Click() 'Ticket #18810
Dim Msg As String, a%
Dim SQLQ  As String
Dim xID As Long
    
    If Not (Data1.Recordset.EOF Or Data1.Recordset.BOF) Then
        Msg = lStr("Are you sure you want to Mark all Follow-Up Records")
        Msg = Msg & Chr(10) & "listed as completed?"
        a% = MsgBox(Msg, 36, "Confirm Update - Mark all completed?")
        If a% <> 6 Then
            Exit Sub
        End If
        xID = Data1.Recordset("EF_FOLLOWUP_ID")
        SQLQ = "UPDATE HR_FOLLOW_UP SET EF_COMPLETED = -1"
        SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
        SQLQ = SQLQ & " AND EF_FREAS = '" & glbFollowUpList & "'"   'SREV, PREV, EDUC
        'If glbFollowUpList = "EDUC" Then
        '    SQLQ = SQLQ & " AND EF_COMMENTS = '" & Replace(frmESEMINARS.txtCourseName, "'", "''") & "'"
        'End If
        gdbAdoIhr001.Execute SQLQ
                
        Data1.Refresh
        Data1.Recordset.Find "EF_FOLLOWUP_ID=" & xID
        
        If Data1.Recordset.EOF Then
            cmdOK.Enabled = False
            cmdCancel.Caption = "Close"
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Dim Msg As String, a%

Call Update_FollowUp_Records

Unload Me

End Sub

Private Sub Form_Load()

glbOnTop = "FRMFOLLOWUPLIST"
'SREV'"  'PREV (Performance), EDUC

Me.Caption = "Mark Older " & IIf(glbFollowUpList = "SREV", "Salary Review ", IIf(glbFollowUpList = "PREV", lStr("Performance Review "), IIf(glbFollowUpList = "EDUC", "Continuing Education ", ""))) & lStr("Follow-ups") & " Records as Completed"

Data1.ConnectionString = glbAdoIHRDBW

If Not glbtermopen Then
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

vbxTrueGrid.Columns(3).Caption = lStr("Administered By")


If Not EERetrieve() Then
    MsgBox "Sorry, Employee can not be found"
    Exit Sub
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If

End Sub

Private Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError

SQLQ = "SELECT * FROM HR_FOLLOW_UP "
SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
SQLQ = SQLQ & " AND EF_FREAS = '" & glbFollowUpList & "'"   'SREV, PREV, EDUC
SQLQ = SQLQ & " AND EF_COMPLETED = 0"  'Not completed
'If glbFollowUpList = "EDUC" Then
'    SQLQ = SQLQ & " AND EF_COMMENTS = '" & Replace(frmESEMINARS.txtCourseName, "'", "''") & "'"
'End If
SQLQ = SQLQ & " ORDER BY EF_FDATE ASC"
        
Data1.RecordSource = SQLQ
Data1.Refresh
If Data1.Recordset.EOF Then cmdOK.Enabled = False
EERetrieve = True

Exit Function
EERError:
End Function

Private Sub lblEEID_Change()

Me.Caption = IIf(glbFollowUpList = "SREV", "Salary Review ", IIf(glbFollowUpList = "PREV", lStr("Performance Review "), IIf(glbFollowUpList = "EDUC", "Continuing Education ", ""))) & lStr("Follow-ups") & " List"

frmFollowUpList.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub

Private Sub Update_FollowUp_Records()
    Dim rsFollowUp As New ADODB.Recordset
    Dim SQLQ As String

    Data1.Recordset.MoveFirst
    With Data1.Recordset
        Do Until .EOF
            
            SQLQ = "SELECT * FROM HR_FOLLOW_UP "
            SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
            SQLQ = SQLQ & " AND EF_FREAS = '" & glbFollowUpList & "'"   'SREV, PREV, EDUC
            SQLQ = SQLQ & " AND EF_FOLLOWUP_ID = " & !EF_FOLLOWUP_ID
            SQLQ = SQLQ & " AND EF_COMPLETED <> " & IIf(!EF_COMPLETED, -1, 0) 'Only if record has changed
            rsFollowUp.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
            If Not rsFollowUp.EOF Then
                rsFollowUp("EF_COMPLETED") = !EF_COMPLETED
                rsFollowUp("EF_LDATE") = Date
                rsFollowUp("EF_LTIME") = Time$
                rsFollowUp("EF_LUSER") = glbUserID
                rsFollowUp.Update
            End If
            rsFollowUp.Close
            Set rsFollowUp = Nothing

        .MoveNext
        Loop
    End With
End Sub

Private Sub vbxTrueGrid_DblClick()
    If Data1.Recordset!EF_COMPLETED Then
        Data1.Recordset!EF_COMPLETED = 0
    Else
        Data1.Recordset!EF_COMPLETED = -1
    End If
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
           
    If vbxTrueGrid.Tag = "ASC" Then
        vbxTrueGrid.Tag = "DESC"
    Else
        vbxTrueGrid.Tag = "ASC"
    End If
    
    SQLQ = "SELECT * FROM HR_FOLLOW_UP "
    SQLQ = SQLQ & " WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = '" & glbFollowUpList & "'"   'SREV, PREV, EDUC
    SQLQ = SQLQ & " AND EF_COMPLETED = 0"  'Not completed
    'If glbFollowUpList = "EDUC" Then
    '    SQLQ = SQLQ & " AND EF_COMMENTS = '" & Replace(frmESEMINARS.txtCourseName, "'", "''") & "'"
    'End If
    SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
    

    Data1.RecordSource = SQLQ
    Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

