VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmPosCopy 
   Caption         =   "Copy Position"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8790
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtToPosDescr 
      Appearance      =   0  'Flat
      DataField       =   "JB_DESCR"
      Height          =   285
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   2
      Tag             =   "00-Enter copy To Position Description"
      Top             =   600
      Width           =   5295
   End
   Begin VB.TextBox txtToPosCode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1870
      MaxLength       =   25
      TabIndex        =   1
      Tag             =   "00-Enter copy To Position Code"
      Top             =   600
      Width           =   1305
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   8790
      _Version        =   65536
      _ExtentX        =   15505
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
         Caption         =   "Cancel"
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
         Left            =   4508
         TabIndex        =   4
         Tag             =   "Cancel changes"
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
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
         Left            =   2708
         TabIndex        =   3
         Tag             =   "Save changes made"
         Top             =   30
         Width           =   1095
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8490
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   3
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin INFOHR_Controls.CodeLookup clpFromJob 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Tag             =   "00-Enter copy From Position Code "
      Top             =   240
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   25
      LookupType      =   5
   End
   Begin VB.Label lblFromPos 
      AutoSize        =   -1  'True
      Caption         =   "From Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   285
      Width           =   1155
   End
   Begin VB.Label lblToPos 
      AutoSize        =   -1  'True
      Caption         =   "To Position"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   645
      Width           =   975
   End
End
Attribute VB_Name = "frmPosCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdCancel_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_Click()
Dim X
Dim SQLQ As String
Dim xFList As String

    If Not CriCheck() Then
        Exit Sub
    End If

    'Create new To Position and update with From Job information
    xFList = Get_Fields(gdbAdoIhr001, "HRJOB", "JB_ID,JB_CODE,JB_DESCR,JB_LDATE,JB_LTIME,JB_LUSER,JB_NBRFIL,JB_FTETOTNU,JB_FTETOTHR,JB_POINTS")
    SQLQ = "INSERT INTO HRJOB (" & xFList & ", JB_CODE, JB_DESCR, JB_LDATE, JB_LTIME, JB_LUSER) "
    SQLQ = SQLQ & in_SQL(glbIHRDB)
    SQLQ = SQLQ & " SELECT  " & xFList & ", '" & txtToPosCode.Text & "' AS JB_CODE, '" & txtToPosDescr.Text & "' AS JB_DESCR "
    SQLQ = SQLQ & ", " & Date_SQL(Date) & " AS JB_LDATE "
    SQLQ = SQLQ & ", '" & Time$ & "' AS JB_LTIME "
    SQLQ = SQLQ & ", '" & glbUserID & "' AS JB_LUSER "
    SQLQ = SQLQ & " FROM HRJOB WHERE (HRJOB.JB_CODE= '" & clpFromJob.Text & "')"
    gdbAdoIhr001.Execute SQLQ
                               
    Unload Me

End Sub

Private Function CriCheck()
Dim xToJobDesc As String

CriCheck = False

If Len(clpFromJob.Text) = 0 Then
    MsgBox "From Position code cannot be blank"
    clpFromJob.SetFocus
    Exit Function
End If

If clpFromJob.Caption = "Unassigned" Then
    MsgBox "Invalid From Position code"
    clpFromJob.SetFocus
    Exit Function
End If

If Len(txtToPosCode.Text) = 0 Then
    MsgBox "To Position code cannot be blank"
    txtToPosCode.SetFocus
    Exit Function
End If
If Len(txtToPosDescr.Text) = 0 Then
    MsgBox "To Position Description cannot be blank"
    txtToPosDescr.SetFocus
    Exit Function
End If

'Check if the New Job Code already exits
xToJobDesc = ""
xToJobDesc = GetJobData(txtToPosCode, "JB_DESCR", "")
If Len(xToJobDesc) > 0 Then
    MsgBox "'To Position' code already exists. Please enter a new 'To Position' code."
    txtToPosCode.SetFocus
    Exit Function
End If

CriCheck = True

End Function

Private Sub clpFromJob_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdOK_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtToPosCode_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtToPosCode_LostFocus()

If Len(txtToPosDescr.Text) = 0 Then
    txtToPosDescr = GetJobData(txtToPosCode, "JB_DESCR", "")
End If

End Sub

Private Sub Form_Load()

Call INI_Controls(Me)

End Sub

Private Sub txtToPosDescr_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
