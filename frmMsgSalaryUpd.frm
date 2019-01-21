VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmMsgSalUpd 
   Caption         =   "Salary Change"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   3345
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   7
      Top             =   2790
      Width           =   4935
      _Version        =   65536
      _ExtentX        =   8705
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
         Left            =   1140
         TabIndex        =   4
         Tag             =   "Save changes made"
         Top             =   30
         Width           =   1095
      End
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
         Left            =   2700
         TabIndex        =   5
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
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LUSER"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5580
      MaxLength       =   25
      TabIndex        =   10
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LTIME"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   5220
      MaxLength       =   25
      TabIndex        =   9
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      DataField       =   "ED_LDATE"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   4950
      MaxLength       =   25
      TabIndex        =   8
      Top             =   5250
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Frame frmBasic 
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   -90
      TabIndex        =   6
      Top             =   0
      Width           =   8235
      Begin VB.ComboBox comPayPer 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1875
         TabIndex        =   3
         Tag             =   "Choose Annum or Hour"
         ToolTipText     =   "Choose Annum or Hour"
         Top             =   2070
         Width           =   1210
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Tag             =   "41-From Date"
         ToolTipText     =   "Attendance From Date"
         Top             =   480
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDateRange 
         Height          =   285
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Tag             =   "41-From Date"
         ToolTipText     =   "Attendance To Date"
         Top             =   855
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin MSMask.MaskEdBox medSalary 
         Height          =   285
         Left            =   1875
         TabIndex        =   2
         Tag             =   "00-Usual working Salary"
         ToolTipText     =   "Salary to update with"
         Top             =   1680
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblSalCode 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "H/A"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2760
         TabIndex        =   17
         Top             =   2160
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Attendance Records:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Update with:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "From Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   14
         Tag             =   "41-Date Terminated"
         Top             =   525
         Width           =   735
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "To Date"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   13
         Tag             =   "41-Date Terminated"
         Top             =   900
         Width           =   585
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   1725
         Width           =   435
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Per"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   11
         Top             =   2130
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmMsgSalUpd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim X%
            
    For X% = 0 To 1
        If Len(dlpDateRange(X%).Text) < 1 Then
            MsgBox ("Both From and To Date are required")
            dlpDateRange(X%).SetFocus
            Exit Sub
        End If
        
        If Len(dlpDateRange(X%).Text) > 0 Then
            If Not IsDate(dlpDateRange(X%).Text) Then
                If X% = 0 Then
                    MsgBox "Invalid From Date"
                Else
                    MsgBox "Invalid To Date"
                End If
                dlpDateRange(X%).Text = ""
                dlpDateRange(X%).SetFocus
                Exit Sub
            End If
        End If
    Next X%
    
    If IsDate(dlpDateRange(0)) And IsDate(dlpDateRange(1)) Then
        If DaysBetween(dlpDateRange(0), dlpDateRange(1)) < 0 Then
            MsgBox "To Date can't be prior to From Date!"
            Me.dlpDateRange(0).SetFocus
            Exit Sub
        End If
    End If
                
    If Len(medSalary) < 1 Then
        MsgBox "Salary is required"
        medSalary.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(medSalary) Then
        MsgBox "Invalid Salary"
        medSalary.SetFocus
        Exit Sub
    End If
    
    If comPayPer.Text = "" Then
        MsgBox "Per cannot be blank"
        comPayPer.SetFocus
        Exit Sub
    End If
    

    'Update Attendance records for the Date Range with the Salary information
    If updAttendanceSalaryInfo Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
MDIMain.panHelp(0).Caption = "info:HR Salary Change"

comPayPer.Clear
comPayPer.AddItem "Annum"
comPayPer.AddItem "Hour"
comPayPer.AddItem "Monthly"

Call setCaption(lblTitle(2))
Call setCaption(lblTitle(3))

'Call INI_Controls(Me)
'lblTitle(1).Caption = lStr(lblTitle(1).Caption)
End Sub

Private Function updAttendanceSalaryInfo() As Boolean
Dim SQLQ As String
Dim rsAttend As New ADODB.Recordset
Dim xRows As Long
Dim xRow As Long

On Error GoTo updAttendanceSalaryInfo_Err

updAttendanceSalaryInfo = False

Screen.MousePointer = vbHourglass
MDIMain.panHelp(0).FloodType = 1
MDIMain.panHelp(1).Caption = "Updating Attendance records..."

SQLQ = "SELECT * FROM HR_ATTENDANCE "
SQLQ = SQLQ & " WHERE AD_EMPNBR = " & glbLEE_ID
SQLQ = SQLQ & " AND AD_DOA >=" & Date_SQL(dlpDateRange(0))
SQLQ = SQLQ & " AND AD_DOA <=" & Date_SQL(dlpDateRange(1))
rsAttend.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
If Not rsAttend.EOF Then
    xRows = rsAttend.RecordCount
    xRow = 1
    
    rsAttend.MoveFirst
    
    Do While Not rsAttend.EOF
        MDIMain.panHelp(0).FloodPercent = (xRow / xRows) * 100
        
        rsAttend("AD_SALARY") = medSalary
        If comPayPer.ListIndex = 0 Then rsAttend("AD_SALCD") = "A"
        If comPayPer.ListIndex = 1 Then rsAttend("AD_SALCD") = "H"
        If comPayPer.ListIndex = 2 Then rsAttend("AD_SALCD") = "M"
        If comPayPer.ListIndex = 3 Then rsAttend("AD_SALCD") = "D"
        
        If rsAttend("AD_UPLOAD") = "Y" Then rsAttend("AD_UPLOAD") = "N"
        
        rsAttend("AD_LDATE") = Date
        rsAttend("AD_LTIME") = Time$
        rsAttend("AD_LUSER") = glbUserID
        rsAttend.Update
        
        xRow = xRow + 1
        rsAttend.MoveNext
    Loop
Else
    Screen.MousePointer = DEFAULT
    MsgBox "There are no Attendance records for the Date Range specified.", vbExclamation, "Attendance Salary Change"
End If
rsAttend.Close
Set rsAttend = Nothing

MDIMain.panHelp(0).FloodPercent = 0
'MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
Screen.MousePointer = DEFAULT

If xRow > 1 Then
    MsgBox xRow - 1 & " Attendance records updated successfully with the Salary Information.", vbInformation, "Attendance Salary Change"
    updAttendanceSalaryInfo = True
End If

Exit Function

updAttendanceSalaryInfo_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Attendance Salary Change", "HR_ATTENDANCE", "updAttendanceSalaryInfo")

End Function

Private Sub medSalary_LostFocus()
If Not IsNumeric(medSalary) Then Exit Sub 'medSalary = 0
If glbFrench Then
    medSalary = Round2DEC(medSalary)    'Val() causing the values to trunc to 0 decimal places
Else
    medSalary = Round2DEC(Val(medSalary))
End If
End Sub

Private Function Round2DEC(tmpNUM, Optional HourlyRate As String)    'laura nov 10, 1997
Dim strNUM As String, X%

If glbFrench Then
    tmpNUM = Replace(Replace(tmpNUM, ",", "."), " ", "")
End If

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
If glbCompSerial = "S/N - 2375W" Then   'City of Timmins
    If GetEmpData(glbLEE_ID, "ED_REGION") <> "S" Then
        Round2DEC = Round(tmpNUM, 2)
    Else
        Round2DEC = Round(tmpNUM, glbCompDecHR)
    End If
Else
    Round2DEC = Round(Val(tmpNUM), glbCompDecHR)
End If
If glbWFC And locCountry = "AUSTRALIA" Then
    If HourlyRate = "Y" Then
        locCompDecHR = 4
    End If
    Round2DEC = Round(tmpNUM, locCompDecHR)
End If
End Function

