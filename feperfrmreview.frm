VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEPERFORMReview 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Staff Profile"
   ClientHeight    =   7575
   ClientLeft      =   405
   ClientTop       =   1365
   ClientWidth     =   11400
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7575
   ScaleWidth      =   11400
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Data2 
      Height          =   375
      Left            =   480
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   979
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
      BevelInner      =   2
      Font3D          =   1
      Alignment       =   1
      Enabled         =   0   'False
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   270
         TabIndex        =   14
         Top             =   180
         Width           =   1005
      End
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
         Left            =   1530
         TabIndex        =   13
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         Left            =   3120
         TabIndex        =   12
         Top             =   135
         Width           =   1740
      End
   End
   Begin INFOHR_Controls.EmployeeLookup elpReptAuthShow 
      DataField       =   "PH_REPTAU"
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   4
      Tag             =   "00-Employee Number of individual's supervisor"
      Top             =   3990
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   503
      ShowUnassigned  =   1
      RefreshDescriptionWhen=   2
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtReptAuthority 
      Appearance      =   0  'Flat
      DataField       =   "PH_REPTAU"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2340
      MaxLength       =   12
      TabIndex        =   24
      Tag             =   "00-Employee Number of individual's supervisor"
      Top             =   3990
      Visible         =   0   'False
      Width           =   870
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "feperfrmreview.frx":0000
      Height          =   1455
      Left            =   0
      OleObjectBlob   =   "feperfrmreview.frx":0014
      TabIndex        =   0
      Top             =   600
      Width           =   9135
   End
   Begin INFOHR_Controls.DateLookup dlpReviewDate 
      DataField       =   "PH_PNEXT"
      Height          =   285
      Index           =   1
      Left            =   2130
      TabIndex        =   5
      Tag             =   "01-Follow Up Date to Review Performance"
      Top             =   4380
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpReviewDate 
      DataField       =   "PH_PREVIEW"
      Height          =   285
      Index           =   0
      Left            =   2130
      TabIndex        =   1
      Tag             =   "01-Event Date"
      Top             =   2850
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PH_CATECODE"
      Height          =   285
      Index           =   1
      Left            =   2130
      TabIndex        =   2
      Tag             =   "00-Performance Category - Code"
      Top             =   3300
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPG"
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   23
      Top             =   6915
      Width           =   11400
      _Version        =   65536
      _ExtentX        =   20108
      _ExtentY        =   1164
      _StockProps     =   15
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8040
         Top             =   120
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
      Begin MSAdodcLib.Adodc Data1 
         Height          =   330
         Left            =   8610
         Top             =   30
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   1
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
   Begin VB.TextBox memComments 
      Appearance      =   0  'Flat
      DataField       =   "PH_COMMENTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   2460
      MaxLength       =   4000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "00-Enter Comments"
      Top             =   5160
      Width           =   6795
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PH_LDATE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   2100
      MaxLength       =   25
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9030
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PH_LTIME"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   3750
      MaxLength       =   25
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9060
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "PH_LUSER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   5490
      MaxLength       =   25
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   9060
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSCheck chkCurrent 
      DataField       =   "PH_CURRENT"
      Height          =   255
      Left            =   7560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "Current Record"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "PH_EVENTCODE"
      Height          =   285
      Index           =   2
      Left            =   2130
      TabIndex        =   3
      Tag             =   "00-Performance Event - Code"
      Top             =   3645
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDPE"
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Performance Event"
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
      Index           =   12
      Left            =   120
      TabIndex        =   25
      Top             =   3660
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Follow-Up Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   90
      TabIndex        =   22
      Top             =   4410
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
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
      Index           =   6
      Left            =   90
      TabIndex        =   21
      Top             =   5190
      Width           =   870
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporting Authority "
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
      Index           =   3
      Left            =   105
      TabIndex        =   20
      Top             =   4020
      Width           =   1650
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Performance Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   105
      TabIndex        =   19
      Top             =   3330
      Width           =   1890
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Event Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   18
      Top             =   2850
      Width           =   1110
   End
   Begin VB.Label lblPerfID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "lblPID"
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
      Height          =   225
      Left            =   5610
      TabIndex        =   17
      Top             =   5340
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "EEId"
      DataField       =   "PH_EMPNBR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   7020
      TabIndex        =   15
      Top             =   5685
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Comp"
      DataField       =   "PH_COMPNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   6930
      TabIndex        =   16
      Top             =   4950
      Visible         =   0   'False
      Width           =   405
   End
End
Attribute VB_Name = "frmEPERFORMReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dynaJobHIS As New ADODB.Recordset
Dim fglbJobDesc$
Dim fglbJobID&
Dim xJobCode
Dim savAuth(3)
Dim DataChanged As Boolean
Dim fglbCurSalary@, fglbCurSalEdate  As Variant
Dim fglbNew As Integer    '

Dim orgReviewDate As String
Dim orgComments As String '
Dim orgPosCode As String
Dim rsDATA As New ADODB.Recordset
Dim fglbJobList
Dim flgloaded As Boolean
Dim MailBody

Private Function chkPerformanceReview()
Dim SQLQ As String, Msg$, dd&, Response%
Dim DgDef As Variant, Title$, DCurPDate As Variant
Dim x%
chkPerformanceReview = False

On Error GoTo chkPerH_Err

If Len(dlpReviewDate(0).Text) = 0 And Len(dlpReviewDate(1).Text) = 0 And Len(clpCode(1).Text) = 0 Then
    Msg$ = "Must enter one of Event Date or"
    Msg$ = Msg$ & Chr(10) & "Follow Up Date or Performance Category"
    DgDef = MB_OKCANCEL '+ MB_ICONQUESTION + MB_DEFBUTTON2
    Response% = MsgBox(Msg$, DgDef, "Warning!")
        If Response% = IDOK Then
            dlpReviewDate(0).SetFocus
            Exit Function
        Else
            Unload Me
            Exit Function
        End If
End If

If Len(dlpReviewDate(0).Text) > 0 Then
    If Not IsDate(dlpReviewDate(0).Text) Then
        Msg$ = "Not a Valid Event Date"
        dlpReviewDate(0).SetFocus
        MsgBox Msg$
        Exit Function
    Else
        If glbSetPer Then
            DCurPDate = CurPDate()
            If DCurPDate > 0 Then    ' 0 if no current record out there
                DCurPDate = CVDate(DCurPDate)
                If DateDiff("d", CVDate(dlpReviewDate(0).Text), DCurPDate) <= 0 Then
                    Msg$ = "Warning...you cannot add or edit a record with a date"
                    Msg$ = Msg$ & Chr(10) & "the same or later than your most current record."
                    Msg$ = Msg$ & Chr(10) & "If you need to edit current performance, "
                    Msg$ = Msg$ & Chr(10) & "go to Performance screen under Employee Menu."
                    MsgBox Msg$
                    dlpReviewDate(0).SetFocus
                    Exit Function
                End If
            End If
        End If
    End If
End If

If Len(dlpReviewDate(1).Text) > 0 Then
    If Not IsDate(dlpReviewDate(1).Text) Then
        Msg$ = "Not a Valid Follow-Up Date"
        dlpReviewDate(1).SetFocus
        MsgBox Msg$
        Exit Function
    Else
        If IsDate(dlpReviewDate(0).Text) Then
            dd& = DateDiff("d", CVDate(dlpReviewDate(0).Text), CVDate(dlpReviewDate(1).Text))
            If dd& < 0 Then
                Msg$ = "Next Review date can not preceed this Review Date."
                dlpReviewDate(1).SetFocus
                MsgBox Msg$
                Exit Function
            End If
        End If
    End If
End If

If fglbNew = True And (Not glbSetPer) Then
    If glbAddHisWarning Then
        DCurPDate = CurPDate()
        If DCurPDate > 0 Then
            DCurPDate = CVDate(DCurPDate)
            If Len(dlpReviewDate(0).Text) > 0 Then
                If DateDiff("d", CVDate(dlpReviewDate(0).Text), DCurPDate) >= 0 Then
                    Msg$ = "Warning, you can not add a record with a date"
                    Msg$ = Msg$ & Chr(10) & "the same or earlier than your most current record."
                    DgDef = MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON2
                    Response% = MsgBox(Msg$)
                  
                    dlpReviewDate(0).SetFocus
                    Exit Function
                  
                End If
            End If
        End If
    End If
End If


If Len(clpCode(1).Text) > 0 Then
    If clpCode(1).Caption = "Unassigned" Then
        MsgBox "If Code Entered Must Be Valid"
        clpCode(1).SetFocus
        Exit Function
    End If
Else
    MsgBox ("Performance Category is Mandatory")
    clpCode(1).SetFocus
    Exit Function
End If

chkPerformanceReview = True

Exit Function

chkPerH_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkPerf", "HR_PERFORM_FRIESEN", "edit/Add")
Call RollBack '28July99 js

End Function


Sub cmdCancel_Click()
Dim x As Integer
On Error GoTo Can_Err
fglbNew = False

rsDATA.CancelUpdate
Call Display_Value

'For x = 0 To 2
    Call txtReptAuthority_Change(0)
'Next
'Call ST_UPD_MODE(True)  ' reset screen's attributes

'fglbNew = False
'Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_PERFORM_FRIESEN", "Cancel")
Call RollBack '28July99 js

End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEPERFORMREVIEW" Then glbOnTop = ""

End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, rc%
Dim x, xID
If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If
On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "These Records?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub

orgReviewDate = dlpReviewDate(1).Text
If orgReviewDate <> "" Then
    If Not updFollow("D") Then
        Exit Sub
    End If
End If


glbSDate = rsDATA("PH_PREVIEW")

If vbxTrueGrid.SelBookmarks.count = 0 Then vbxTrueGrid.SelBookmarks.Add Data1.Recordset.Bookmark
For x = 0 To vbxTrueGrid.SelBookmarks.count - 1
    Data1.Recordset.Bookmark = vbxTrueGrid.SelBookmarks(x)
    xID = Data1.Recordset("PH_ID")
    gdbAdoIhr001.BeginTrans
    gdbAdoIhr001.Execute "DELETE FROM HR_PERFORM_FRIESEN WHERE PH_ID=" & xID
    gdbAdoIhr001.CommitTrans
    DoEvents

Next

Data1.Refresh

If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    Call Set_Current_Flag
Else
    Call Display_Value
End If
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_PERFORM_FRIESEN", "Delete")
Call RollBack '28July99 js

End Sub

'Private Sub cmdDelete_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdModify_Click()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String
Dim Response%, Msg$, Title$, DgDef As Double


On Error GoTo Mod_Err

orgReviewDate = dlpReviewDate(1).Text
orgComments = memComments


Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HR_PERFORM_FRIESEN", "Modify")
Call RollBack '28July99 js

End Sub

'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String, Msg$
Dim x

On Error GoTo AddN_Err

'Ticket #21811
If Not IsNumeric(glbUserID) Then
    Msg$ = "To add a new Staff Profile, the User's 'User ID' must be same as his/her Employee #. The Reporting Authority "
    Msg$ = Msg$ & "value cannot be alphanumeric."
    MsgBox Msg$, vbCritical, "Invalid Reporting Authority"
    Exit Sub
End If

fglbNew = True

If Not Set_Position("", True) Then
    Msg$ = "No current position found "
    Msg$ = Msg$ & Chr(10) & "Please review Position prior to updating Performance."
    MsgBox Msg$
    Exit Sub
End If

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

Call Set_Control("B", Me)

rsDATA.AddNew
lblEEID = glbLEE_ID

lblCNum.Caption = "001"
'elpReptAuthShow(0).Text = ShowEmpnbr(savAuth(0))
dlpReviewDate(1).Text = ""
dlpReviewDate(0).SetFocus
'Sam changed it to default as Jerry sent an email to modify
'Please modify the Friesen’s Performance screen to default the Reporting Authority to equal the login User id.Jerry
elpReptAuthShow(0).Text = glbUserID
elpReptAuthShow(0).Enabled = False
clpCode(2).Enabled = True

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_PERFORM_FRIESEN", "Add")
Call RollBack '28July99 js

End Sub

Sub cmdOK_Click()
Dim rsPER As New ADODB.Recordset
Dim x, xReptAuthority, xID
On Error GoTo Add_Err

If Not chkPerformanceReview() Then Exit Sub

Screen.MousePointer = HOURGLASS

If gsEMAIL_ONPERFORMANCE Then
    MailBody = ""
    If NewHireForms.count = 0 Then 'Non new hire
        If fglbNew Or chkCurrent Then
            MailBody = "The " & lStr("Performance") & " has been changed." & vbCrLf & vbCrLf
            MailBody = MailBody & "Employee #: " & lblEENum.Caption & vbCrLf
            MailBody = MailBody & "Name: " & lblEEName.Caption & vbCrLf
            MailBody = MailBody & lblTitle(1) & ": " & dlpReviewDate(0) & vbCrLf
            MailBody = MailBody & lblTitle(5) & ": " & dlpReviewDate(1) & vbCrLf
        End If
    End If
End If

Call UpdUStats(Me) ' update user's stats (who did it and when)
Call Set_Control("U", Me, rsDATA)
    xReptAuthority = getEmpnbr(elpReptAuthShow(0).Text)
    If Not IsNull(xReptAuthority) And xReptAuthority <> "" Then
        rsDATA("PH_REPTAU") = xReptAuthority
    End If
If elpReptAuthShow(0).Caption <> "Unassigned" And elpReptAuthShow(0).Caption <> "" Then
    rsDATA("PH_SUPERNAME") = elpReptAuthShow(0).Caption
End If
rsDATA("PH_EMPNAME") = lblEEName.Caption

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
    xID = rsDATA!PH_ID
    glbSDate = rsDATA("PH_PREVIEW")
    rsDATA.Requery
   
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
    xID = rsDATA!PH_ID
    glbSDate = rsDATA("PH_PREVIEW")
    rsDATA.Requery
     
End If
'Data1.Refresh

Call Set_Current_Flag

Data1.Refresh
Data1.Recordset.Find "PH_ID=" & xID


Call Display_Value

If chkCurrent Then
    If Not updFollow("U") Then Exit Sub
End If


fglbNew = False

Call SET_UP_MODE
'Call ST_UPD_MODE(True)

If gsEMAIL_ONPERFORMANCE Then
    If Len(MailBody) > 0 Then
        Screen.MousePointer = DEFAULT
        Call imgEmail_Click
    End If
End If


Screen.MousePointer = DEFAULT
Call NextForm

Exit Sub

Add_Err:
If Err = 3021 Then
    Err = 0
    Resume Next
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_PERFORM_FRIESEN", "Update")
Call RollBack

End Sub

Private Sub cmdPosCode_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

'Private Sub cmdPosition_Click()
'Unload frmEPOSITION
'glbSetPos = glbSetPer
'frmEPOSITION.Show
'Unload Me

'End Sub

Sub cmdPrint_Click()
Dim RHeading As String, xReport, dscGroup$

'cmdPrint.Enabled = False
RHeading = lblEEName.Caption & "'s Performance History"
Me.vbxCrystal.WindowTitle = RHeading & " Report"

If Not glbtermopen Then
    xReport = glbIHRREPORTS & "rgridprr.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    Me.vbxCrystal.SelectionFormula = "{HR_PERFORM_FRIESEN.PH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    xReport = glbIHRREPORTS & "rgridpr1.rpt"

    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
    End If
        Me.vbxCrystal.SelectionFormula = "{Term_PERFORM_HISTORY.TERM_SEQ}=" & glbTERM_Seq & " "
End If

Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, dscGroup$

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

'cmdPrint.Enabled = False
RHeading = lblEEName.Caption & "'s Performance Information"
Me.vbxCrystal.WindowTitle = RHeading & " Report"

If Not glbtermopen Then
    xReport = glbIHRREPORTS & "rgridprr.rpt"
    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRDB
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRDB
    End If
    Me.vbxCrystal.SelectionFormula = "{HR_PERFORM_FRIESEN.PH_EMPNBR}=" & glbLEE_ID & " "
End If

If glbtermopen Then
    xReport = glbIHRREPORTS & "rgridpr1.rpt"

    Me.vbxCrystal.ReportFileName = xReport
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    Else
        Me.vbxCrystal.Connect = "PWD=petman;"
        Me.vbxCrystal.DataFiles(0) = glbIHRAUDIT
        Me.vbxCrystal.DataFiles(1) = glbIHRDB
        Me.vbxCrystal.DataFiles(2) = glbIHRAUDIT
    End If
        Me.vbxCrystal.SelectionFormula = "{Term_PERFORM_FRIESEN.TERM_SEQ}=" & glbTERM_Seq & " "
End If

Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
'cmdPrint.Enabled = True
End Sub


'Private Sub cmdPrint_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdSalary_Click()
'Unload frmESALARY
'glbSetSal = glbSetPer
'frmESALARY.Show
'Unload Me
'End Sub

'Private Sub CR_Job_Snap()
'Dim SQLQ As String, countr As Integer
'Dim Desc As String
'Dim Msg As String
'
'On Error GoTo Job_Err
'
'Screen.MousePointer = HOURGLASS
'
'SQLQ = "SELECT * FROM HRJOB"
'
'If Job_Snaps.State <> 0 Then Job_Snaps.Close
'Job_Snaps.Open SQLQ, gdbAdoIhr001, adOpenStatic
'
'If Job_Snaps.EOF And Job_Snaps.BOF Then
'    Msg = "No Job descriptions found" & Chr(10)
'    Msg = Msg & "You will require authority to add one to continue"
'    MsgBox Msg
'Else
'    'EOF?
'    Job_Snaps.MoveFirst
'End If
'
'Screen.MousePointer = DEFAULT
'
'Exit Sub
'
'Job_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "List Jobs", "HRJOB", "SELECT")
'Call RollBack '28July99 js
'
'End Sub

Private Function CurPDate()
Dim SQLQ As String
Dim HRP_Snap As New ADODB.Recordset

CurPDate = 0    ' returns 0 if no found records

On Error GoTo JP_Err

SQLQ = "Select HR_PERFORM_FRIESEN.* from HR_PERFORM_FRIESEN"
SQLQ = SQLQ & " where HR_PERFORM_FRIESEN.PH_EMPNBR = " & glbLEE_ID & " "
SQLQ = SQLQ & " AND HR_PERFORM_FRIESEN.PH_CURRENT <>0"

HRP_Snap.Open SQLQ, gdbAdoIhr001, adOpenStatic

If HRP_Snap.BOF And HRP_Snap.EOF Then
    Exit Function
Else
    CurPDate = HRP_Snap("PH_PREVIEW")
    HRP_Snap.Close
End If

Exit Function

JP_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Perform History Snap", "HR_PERFORM_FRIESEN", "SELECT")
Call RollBack '28July99 js

End Function

Function EERetrieve()
Dim SQLQ As String
Dim x, xFld
EERetrieve = False

On Error GoTo EERError

    Screen.MousePointer = HOURGLASS
    
    
    If glbtermopen Then
        SQLQ = "SELECT Term_PERFORM_FRIESEN.*,"
    Else
        SQLQ = "SELECT HR_PERFORM_FRIESEN.*,"
    End If

    If glbOracle Then
        SQLQ = SQLQ & "PH_REPTAU AS REPTAU"
    Else
        SQLQ = SQLQ & "STR(PH_REPTAU) AS REPTAU"
    End If

If glbtermopen Then
    SQLQ = SQLQ & " FROM Term_PERFORM_FRIESEN"
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
Else
    SQLQ = SQLQ & " FROM  HR_PERFORM_FRIESEN"
    SQLQ = SQLQ & " WHERE PH_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC, PH_PNEXT DESC"
    
Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True  'new

Screen.MousePointer = DEFAULT
Call Display_Value

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Staff Profile", "HR_PERFORM_FRIESEN", "SELECT")
Call RollBack

Exit Function

End Function

Private Sub clpCode_DblClick(Index As Integer)
DataChanged = True
End Sub

Private Sub clpCode_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
    If Index = 1 Then   'Performance Category
        clpCode(Index).TransDiv = GetTransDiv(xJobCode)
    End If
End Sub

Private Sub clpCode_KeyPress(Index As Integer, KeyAscii As Integer)
DataChanged = True
End Sub

Private Sub dlpReviewDate_DblClick(Index As Integer)
DataChanged = True
End Sub

Private Sub dlpReviewDate_KeyPress(Index As Integer, KeyAscii As Integer)
DataChanged = True
End Sub

Private Sub elpReptAuthShow_Change(Index As Integer)
txtReptAuthority(Index).Text = getEmpnbr(elpReptAuthShow(Index).Text)
End Sub

Private Sub elpReptAuthShow_DblClick(Index As Integer)
DataChanged = True
End Sub

Private Sub Form_Activate()
glbOnTop = "FRMEPERFORMREVIEW"
'clpPosCode.seleEMPCode = fglbJobList
flgloaded = True
Call SET_UP_MODE

End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMEPERFORMREVIEW"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%
DataChanged = False

glbOnTop = "FRMEPERFORMREVIEW"

If glbtermopen Then
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
    If glbNoNONE Then
        If glbUNION = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbNoEXEC Then        'Hemu -EXE
        If glbUNION = "EXEC" Then   'Hemu -EXE
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
    If glbNoNONE Then
        If glbUNIONTe = "NONE" Then
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
    If glbNoEXEC Then        'Hemu -EXE
        If glbUNIONTe = "EXEC" Then     'Hemu -EXE
            MsgBox "You Do Not Have Authority For This Transaction"
            glbOnTop = Empty
            Unload Me
            Screen.MousePointer = DEFAULT
            Exit Sub
        End If
    End If
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If
Screen.MousePointer = HOURGLASS
If Len(glbLEE_SName) < 1 Then Exit Sub
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = IIf(glbSetPer, "Set ", "") & "Staff Profile - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)

Call CR_JobHis_Snap
'Call CR_SalHis_Snap
If Data1.Recordset.EOF Then
    If Not Set_Position("", True) Then Exit Sub
'    If Set_Salary("", True) Then
'        lblCSalary = Round2DEC(fglbCurSalary@)
'        lblEDateD = fglbCurSalEdate
'    End If
End If
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Call Display_Value
If Not gSec_Upd_Performance Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If
'fraCurrentSalary.Visible = gSec_Inq_Salary
Call INI_Controls(Me)
Screen.MousePointer = DEFAULT
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Dim VR
'If DataChanged = True Then
'        VR = MsgBox("Do you want to save changes?", MB_YESNO)
'        If VR = IDYES Then
'            Me.cmdOK_Click
'        End If
'End If


Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)

End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmEPERFORMReview = Nothing  'carmen may 00
    Call NextForm
End Sub

'Private Function getJOB()
'Dim SQLQ As String
'Dim rsJOB As New ADODB.Recordset
'getJOB = False
'On Error GoTo Jobd_Err
'If Len(lblJOB) > 0 Then
'    SQLQ = "SELECT * FROM HRJOB WHERE JB_CODE = '" & lblJOB & "'"
'    rsJOB.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
'    If rsJOB.EOF Then
'        clpCode(1).Caption = "Unassigned"
'        Exit Function
'    End If
'    getJOB = True
'    clpCode(1).Caption = rsJOB("JB_DESCR")
'    rsJOB.Close
'End If
'
'
'Exit Function
'
'Jobd_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Job Snap", "HRJOB", "SELECT")
'Call RollBack '28July99 js
'
'End Function


Private Sub medSalary_GotFocus()
 Call SetPanHelp(ActiveControl)
End Sub

Private Sub imgNoSec_Click()

End Sub

Private Sub memComments_GotFocus()
    Call SetPanHelp(ActiveControl)
    
End Sub

Private Sub Set_Current_Flag()
Dim SQLQ As String, Msg$
Dim dyn_HRPHHIS As New ADODB.Recordset

On Error GoTo SCFError
If glbMulti Then Exit Sub

'Hemu - 07/07/2003 Begin - Commented out the clone line cause it was giving Error
'                          as 'Row cannot be located for updating'
'Set dyn_HRPHHIS = Data1.Recordset.Clone
dyn_HRPHHIS.Open Data1.RecordSource, gdbAdoIhr001, adOpenStatic, adLockOptimistic
'Hemu- 07/07/2003  End

Screen.MousePointer = HOURGLASS

If dyn_HRPHHIS.RecordCount < 1 Then
    Screen.MousePointer = DEFAULT
    Exit Sub
End If

'Hemu - 07/07/2003 Begin -Check # of records before moving first
If dyn_HRPHHIS.RecordCount > 0 Then dyn_HRPHHIS.MoveFirst
'Hemu - 07/07/2003 End

'EOF?
dyn_HRPHHIS("PH_CURRENT") = True
dyn_HRPHHIS.Update
dyn_HRPHHIS.MoveNext

While Not dyn_HRPHHIS.EOF
    'Hemu - 07/07/2003 Begin - to improve speed, Jaddy suggested
    If dyn_HRPHHIS("PH_CURRENT") <> 0 Then
        dyn_HRPHHIS("PH_CURRENT") = False
        dyn_HRPHHIS.Update
    End If
    'Hemu - 07/07/2003 End
    dyn_HRPHHIS.MoveNext
Wend

dyn_HRPHHIS.Close
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Screen.MousePointer = DEFAULT

Exit Sub

SCFError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_PERFORM_FRIESEN", "Add")
Call RollBack '28July99 js

End Sub

Private Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If
chkCurrent.Enabled = TF
memComments.Enabled = TF
clpCode(1).Enabled = TF
dlpReviewDate(0).Enabled = TF
dlpReviewDate(1).Enabled = TF
'elpReptAuthShow(0).Enabled = TF
    Me.cmdModify_Click

glbDocName = "Performance"
If gsAttachment_DB Then
   
    Call DispimgIcon(Me, "frmEPERFORMReview")
End If


End Sub

Private Sub memComments_KeyPress(KeyAscii As Integer)
DataChanged = True
End Sub


Private Sub txtReptAuthority_Change(Index As Integer)
    elpReptAuthShow(Index).Text = ShowEmpnbr(txtReptAuthority(Index).Text)
End Sub


Private Function updFollow(xType)
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim Edit1 As Integer

updFollow = False

On Error GoTo CrFollow_Err

If orgReviewDate <> "" Then  ' DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'PREV'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(orgReviewDate)
   
    dynHRAT.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If dynHRAT.BOF And dynHRAT.EOF Then
        Edit1 = False
    Else
        Edit1 = True    ' returns true if found records
    End If
Else
    Edit1 = False
End If

If xType = "U" Then
    
    rsTB.Open "HR_FOLLOW_UP", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdTableDirect
    If fglbNew And dlpReviewDate(1).Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpReviewDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "PREV"
           
        rsTB("EF_COMMENTS") = memComments
              
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And dlpReviewDate(1).Text <> "" Then
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpReviewDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "PREV"
                
        rsTB("EF_COMMENTS") = memComments
                
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpReviewDate(1).Text <> "" Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(dlpReviewDate(1).Text)
            dynHRAT("EF_FREAS") = "PREV"
                        
            dynHRAT("EF_COMMENTS") = memComments
                        
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If orgComments <> memComments Or orgReviewDate <> dlpReviewDate(1).Text Then
            Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpReviewDate(1).Text = "" Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    End If
Else
    If Edit1 = True Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
        Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If dlpReviewDate(1).Text = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "Check the Follow up table"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

End Function
'Private Sub txtReviewDate_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub vbxTrueGrid_BeforeRowColChange(Cancel As Integer)
Cancel = Not isUpdated(Me)
End Sub

Private Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim tdcode$, x As Integer
Dim SQLQ As String

On Error GoTo Tab1_Err
Call Display_Value

'If Not Data1.Recordset.EOF Then
'
'
'   ' If Not Set_Position(clpPosCode.Text, False) Then Exit Sub
'
''    If Set_Salary(clpPosCode.Text, True) Then
''        lblCSalary = Round2DEC(fglbCurSalary@)
''        lblEDateD = fglbCurSalEdate
''    Else
''        lblCSalary = ""
''        lblEDateD = ""
''    End If
'Else
'    'lblJobDesc.Caption = "Unassigned"
'   ' lblCSalary = ""
'   ' lblEDateD = ""
'End If

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HR_PERFORM_FRIESEN", "Add")
Resume Next

End Sub

Private Function RollBack()
On Error GoTo rr
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
rr:
End Function

Private Sub CR_JobHis_Snap()
Dim SQLQ As String, countr As Integer
Dim Desc As String
Dim Msg As String

On Error GoTo JobHis_Err

Screen.MousePointer = HOURGLASS
If glbtermopen Then
    SQLQ = "Select * from Term_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
Else
    SQLQ = "Select * from HR_JOB_HISTORY "
    SQLQ = SQLQ & " WHERE JH_EMPNBR=" & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY JH_CURRENT " & IIf(glbSQL, "DESC", "") & ",JH_SDATE DESC"

    If dynaJobHIS.State <> 0 Then dynaJobHIS.Close
    dynaJobHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
End If

If Not dynaJobHIS.EOF Then
    Do Until dynaJobHIS.EOF
        If Not IsNull(dynaJobHIS!JH_JOB) Then
            fglbJobList = fglbJobList & dynaJobHIS!JH_JOB & ","
        End If
        dynaJobHIS.MoveNext
    Loop
    If Right(fglbJobList, 1) = "," Then
        fglbJobList = Left(fglbJobList, Len(fglbJobList) - 1)
    End If
    dynaJobHIS.MoveFirst
End If
Screen.MousePointer = DEFAULT

Exit Sub

JobHis_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hours per Week", "HR_JOB_History", "SELECT")
Screen.MousePointer = DEFAULT
Resume Next

End Sub

'Private Sub CR_SalHis_Snap()
'Dim SQLQ As String, countr As Integer
'Dim Desc As String
'Dim Msg As String
'
'On Error GoTo SalHis_Err
'
'Screen.MousePointer = HOURGLASS
'If glbtermopen Then
'    SQLQ = "Select * from Term_SALARY_HISTORY "
'    SQLQ = SQLQ & " WHERE TERM_SEQ=" & glbTERM_Seq
'    SQLQ = SQLQ & " ORDER BY SH_CURRENT " & IIf(glbSQL, "DESC", "") & ",SH_SDATE DESC"
'
'    If dynaSalHIS.State <> 0 Then dynaSalHIS.Close
'    dynaSalHIS.Open SQLQ, gdbAdoIhr001X, adOpenStatic
'Else
'    SQLQ = "Select * from HR_SALARY_HISTORY "
'    SQLQ = SQLQ & " WHERE SH_EMPNBR=" & glbLEE_ID
'    SQLQ = SQLQ & " ORDER BY SH_CURRENT " & IIf(glbSQL, "DESC", "") & ",SH_SDATE DESC"
'
'    If dynaSalHIS.State <> 0 Then dynaSalHIS.Close
'    dynaSalHIS.Open SQLQ, gdbAdoIhr001, adOpenStatic
'End If
'Screen.MousePointer = DEFAULT
'
'Exit Sub
'
'SalHis_Err:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Hours per Week", "HR_SALARY_HISTORY", "SELECT")
'Screen.MousePointer = DEFAULT
'Resume Next
'
'End Sub


Private Function Set_Position(nJob As String, nCurrent As Boolean)
Dim SQLQ As String, Msg$

Set_Position = False
On Error GoTo SCError
Screen.MousePointer = HOURGLASS
dynaJobHIS.Requery

SQLQ = ""
If nCurrent Then SQLQ = SQLQ & " JH_CURRENT<>0 "
If nJob <> "" Then SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " JH_JOB='" & nJob & "' "
dynaJobHIS.Filter = SQLQ

If dynaJobHIS.BOF And dynaJobHIS.EOF Then
    glbStopPerform% = nCurrent
    Screen.MousePointer = DEFAULT
    dynaJobHIS.Filter = ""
    Exit Function
Else
    glbStopPerform% = False
End If

xJobCode = dynaJobHIS("JH_JOB")      ' record
'If IsNull(dynaJobHIS("JH_ID")) Then fglbJobID& = 0 Else fglbJobID& = dynaJobHIS("JH_ID")
'If IsNull(dynaJobHIS("JH_REPTAU")) Then savAuth(0) = "" Else savAuth(0) = dynaJobHIS("JH_REPTAU")
'If IsNull(dynaJobHIS("JH_REPTAU2")) Then savAuth(1) = "" Else savAuth(1) = dynaJobHIS("JH_REPTAU2")
'If IsNull(dynaJobHIS("JH_REPTAU3")) Then savAuth(2) = "" Else savAuth(2) = dynaJobHIS("JH_REPTAU3")
dynaJobHIS.Filter = ""
Set_Position = True

Screen.MousePointer = DEFAULT
Exit Function

SCError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_JOB_HISTORY", "SELECT")
Call RollBack '28July99 js

End Function


'Private Function Set_Salary(nJob As String, nCurrent As Boolean)
'Dim SQLQ As String, Msg$
'
'Set_Salary = False
'On Error GoTo SCError
'Screen.MousePointer = HOURGLASS
'dynaSalHIS.Requery
'SQLQ = ""
'If nCurrent Then SQLQ = SQLQ & " SH_CURRENT<>0 "
'If nJob <> "" Then SQLQ = SQLQ & IIf(SQLQ = "", "", "AND") & " SH_JOB='" & nJob & "' "
'dynaSalHIS.Filter = SQLQ
'
'If dynaSalHIS.BOF And dynaSalHIS.EOF Then
'    glbStopPerform% = nCurrent
'    Screen.MousePointer = DEFAULT
'    dynaSalHIS.Filter = ""
'    Exit Function
'Else
'    glbStopPerform% = False
'End If
'
'fglbCurSalary@ = dynaSalHIS("SH_SALARY")
'fglbCurSalEdate = dynaSalHIS("SH_EDATE")
'
'dynaSalHIS.Filter = ""
'Set_Salary = True
'Screen.MousePointer = DEFAULT
'Exit Function
'
'SCError:
'glbFrmCaption$ = Me.Caption
'glbErrNum& = Err
'
'Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HR_Salary_HISTORY", "SELECT")
'Call RollBack '28July99 js
'
'End Function

Private Function Round2DEC(tmpNUM) 'laura nov 10, 1997
Dim strNUM As String, x%

If glbCompDecHR <> 2 And glbCompDecHR <> 3 And glbCompDecHR <> 4 Then
    glbCompDecHR = 2  'THIS SHOULD NOT HAPPEN BUT IS A VALID DEFAULT
End If
strNUM = "0." & String(glbCompDecHR, "0")
Round2DEC = Format(Round(tmpNUM, glbCompDecHR), strNUM)

End Function

Sub Display_Value()
    Dim SQLQ
    Dim x, xFld
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        If glbtermopen Then
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        Else
            rsDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        End If
    Else
        If glbtermopen Then
            SQLQ = "SELECT Term_PERFORM_FRIESEN.*,"
        Else
            SQLQ = "SELECT HR_PERFORM_FRIESEN.*,"
        End If
           
        If glbOracle Then
            SQLQ = SQLQ & "PH_REPTAU AS REPTAU"
        Else
            SQLQ = SQLQ & "STR(PH_REPTAU) AS REPTAU"
        End If
        
        If glbtermopen Then
            SQLQ = SQLQ & " FROM Term_PERFORM_FRIESEN "
            SQLQ = SQLQ & " WHERE PH_ID=" & Data1.Recordset!PH_ID
            SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC"
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
        
        Else
            SQLQ = SQLQ & " FROM  HR_PERFORM_FRIESEN"
            SQLQ = SQLQ & " WHERE PH_ID = " & Data1.Recordset!PH_ID
            SQLQ = SQLQ & " ORDER BY PH_PREVIEW DESC"
            If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
            rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        
        
        End If
        
'        If gsAttachment_DB Then
'            If rsDATA.EOF Or rsDATA.BOF Then
'                imgSec.Visible = False
'                imgNoSec.Visible = True
'            End If
'        End If

        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    
        Call Set_Control("R", Me, rsDATA)
        If Not (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
            chkCurrent = Data1.Recordset("PH_CURRENT")
        End If
    End If
    
    Call SET_UP_MODE
    
    Me.cmdModify_Click
    
    If elpReptAuthShow(0).Text = glbUserID Then
        dlpReviewDate(0).Enabled = True
        clpCode(1).Enabled = True
        clpCode(2).Enabled = True
        elpReptAuthShow(0).Enabled = False
        dlpReviewDate(1).Enabled = True
        memComments.Enabled = True
    Else
        elpReptAuthShow(0).Enabled = False
        
        dlpReviewDate(0).Enabled = False
        clpCode(1).Enabled = False
        clpCode(2).Enabled = False
        dlpReviewDate(1).Enabled = False
        memComments.Enabled = False
        
    End If

End Sub

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

Public Property Get RelateMode() As RelateModeEnum
RelateMode = RelateEMP
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = gSec_Upd_Performance
End Property

Public Property Get Addable() As Boolean
Addable = Not glbtermopen
End Property
Public Property Get Updateble() As Boolean

Updateble = Not glbtermopen
End Property
Public Property Get Deleteble() As Boolean

Deleteble = Not glbtermopen
End Property
Public Property Get Printable() As Boolean
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf rsDATA.EOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
If Not Updateble Then TF = False
Call ST_UPD_MODE(TF)
End Sub

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmEPERFORMReview.Caption = "Staff Profile - " & Left$(glbLEE_SName, 5)
    frmEPERFORMReview.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
 If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
End Sub

Private Function GetTransDiv(xPos)
Dim rsTran As New ADODB.Recordset
Dim SQLQ As String
Dim xPosGroup As String
Dim xPerfCategort As String

    xPosGroup = GetJobData(xPos, "JB_GRPCD")
    xPerfCategort = "'*'"
    If Not Len(xPosGroup) = 0 Then
        SQLQ = "SELECT * FROM HR_PERF_JOBGRP "
        SQLQ = SQLQ & "WHERE PJ_GRPCD ='" & xPosGroup & "' "
        rsTran.Open SQLQ, gdbAdoIhr001, adOpenStatic
        Do While Not rsTran.EOF
            xPerfCategort = xPerfCategort & ",'" & rsTran("PJ_CATECODE") & "'"
            rsTran.MoveNext
        Loop
    End If
    GetTransDiv = xPerfCategort
End Function

Public Sub imgEmail_Click()
Dim xEmail
Dim xToEmail As String
On Error GoTo Email_Err
    If gsEMAIL_ONPERFORMANCE Then
        If Not UserEmailExist Then
            Exit Sub
        End If
        'xEmail = GetCurEmpEmail
        'xEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE")
            
        'Ticket #20317 - Send email to More Emails list as well.
        xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE", glbLEE_ID)
        If Len(xToEmail) = 0 Then 'cannot find email in More Emails then check Company Preference email
            xToEmail = GetComPreferEmail("EMAIL_ONPERFORMANCE")
        End If
            
        'If Len(xEmail) > 0 Then    'Hemu - (Ticket #11562) - Jerry asked to remove the check for email address presence.
            frmSendEmail.txtTo.Text = xToEmail
            If glbCompSerial = "S/N - 2382W" Then  'Samuel Ticket #18352, do not cc it to employee
            Else
                frmSendEmail.txtCC.Text = GetCurEmpEmail 'xEmail
            End If
            'Ticket #18578
            frmSendEmail.txtSubject.Text = "info:HR " & lStr("Performance") & " Change Notice - " & lblEEName.Caption
            frmSendEmail.txtBody.Text = MailBody
            frmSendEmail.Show 1
        'Else
            'If Len(glbLEE_SName) = 0 Then
            '    MsgBox "There is no email on Status/Dates screen for employee. "
            'Else
            '    MsgBox "There is no email on Status/Dates screen for employee " & glbLEE_SName & ", " & glbLEE_FName & ". "
            'End If
        '    MsgBox "There is no email address for the 'Email Notification on " & lstr("Performance") & " ' on Company Preference screen. "
        'End If


    End If
    Exit Sub

Email_Err:
    If Err.Number = 364 Then
        Exit Sub
    End If
    
    glbFrmCaption$ = Me.Caption
    glbErrNum& = Err
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Send EMail", "SMTP", "SENDEMAIL")
    Resume Next

End Sub

