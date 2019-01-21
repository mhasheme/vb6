VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmCobra 
   Caption         =   "COBRA"
   ClientHeight    =   8025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   11355
   WindowState     =   2  'Maximized
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "END_DATE"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   7
      Left            =   8340
      TabIndex        =   41
      Tag             =   "40-COBRA End Date"
      Top             =   5220
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "PAYMENT_DATE"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   6
      Left            =   8340
      TabIndex        =   40
      Tag             =   "40-Payment Due Date"
      Top             =   4800
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "DATE_PAID"
      Height          =   255
      Index           =   5
      Left            =   8340
      TabIndex        =   39
      Tag             =   "40-Date Paid To"
      Top             =   4410
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "ELECTION_DATE"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   4
      Left            =   8340
      TabIndex        =   38
      Tag             =   "40-Coverage Election Date"
      Top             =   4020
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "REPLY_DATE"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   3
      Left            =   8340
      TabIndex        =   37
      Tag             =   "40-Reply Date"
      Top             =   3120
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "NOTE_DATE"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   3
      Tag             =   "41-Employer Notification Date"
      Top             =   3600
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpDate 
      DataField       =   "EVENT_DATE"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   2
      Tag             =   "41-Date of Qualifying Event"
      Top             =   3180
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   450
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "EVENT"
      DataSource      =   "Data1"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   1
      Tag             =   "01-Qualifying Reason Code"
      Top             =   2760
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "COBR"
      Object.Height          =   255
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   6240
      Top             =   7440
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      Caption         =   "adodc1"
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
   Begin VB.TextBox txtLetter 
      Appearance      =   0  'Flat
      DataField       =   "LETTER_DIR5"
      DataSource      =   "Data1"
      Height          =   345
      Index           =   5
      Left            =   3080
      TabIndex        =   8
      Tag             =   "00-Employee Notification Letter"
      Top             =   5760
      Width           =   3285
   End
   Begin VB.TextBox txtLetter 
      Appearance      =   0  'Flat
      DataField       =   "LETTER_DIR4"
      DataSource      =   "Data1"
      Height          =   345
      Index           =   4
      Left            =   3080
      TabIndex        =   7
      Tag             =   "00-Employee Notification Letter"
      Top             =   5340
      Width           =   3285
   End
   Begin VB.TextBox txtLetter 
      Appearance      =   0  'Flat
      DataField       =   "LETTER_DIR3"
      DataSource      =   "Data1"
      Height          =   345
      Index           =   3
      Left            =   3080
      TabIndex        =   6
      Tag             =   "00-Employee Notification Letter"
      Top             =   4890
      Width           =   3285
   End
   Begin VB.TextBox txtLetter 
      Appearance      =   0  'Flat
      DataField       =   "LETTER_DIR2"
      DataSource      =   "Data1"
      Height          =   345
      Index           =   2
      Left            =   3080
      TabIndex        =   5
      Tag             =   "00-Employee Notification Letter"
      Top             =   4440
      Width           =   3285
   End
   Begin VB.TextBox txtLetter 
      Appearance      =   0  'Flat
      DataField       =   "LETTER_DIR1"
      DataSource      =   "Data1"
      Height          =   345
      Index           =   1
      Left            =   3080
      TabIndex        =   4
      Tag             =   "00-Employee Notification Letter"
      Top             =   3990
      Width           =   3285
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   29
      Top             =   7365
      Width           =   11355
      _Version        =   65536
      _ExtentX        =   20029
      _ExtentY        =   1164
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
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
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
         Left            =   5340
         TabIndex        =   36
         Tag             =   "Print a Dependent Listing Report"
         Top             =   30
         Width           =   855
      End
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
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
         Left            =   4500
         TabIndex        =   35
         Tag             =   "Delete the Dependent Selected"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
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
         Left            =   3660
         TabIndex        =   34
         Tag             =   "Add a new Dependent"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
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
         Left            =   2715
         TabIndex        =   33
         Tag             =   "Cancel the changes made"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
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
         Left            =   1890
         TabIndex        =   32
         Tag             =   "Save the changes made"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
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
         Left            =   1050
         TabIndex        =   31
         Tag             =   "Edit the information on this screen"
         Top             =   30
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
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
         Left            =   180
         TabIndex        =   30
         Tag             =   "Close and exit this screen"
         Top             =   30
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   8730
         Top             =   105
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         ReportSource    =   1
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LUSER"
      DataSource      =   "Data1"
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   2
      Left            =   8295
      TabIndex        =   25
      Text            =   "LSER"
      Top             =   7815
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LTIME"
      DataSource      =   "Data1"
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   1
      Left            =   6975
      TabIndex        =   24
      Text            =   "LTIME"
      Top             =   7815
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "LDATE"
      DataSource      =   "Data1"
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   5700
      TabIndex        =   23
      Text            =   "LDATE"
      Top             =   7815
      Visible         =   0   'False
      Width           =   1215
   End
   Begin Threed.SSCheck ChkCoverCont 
      DataField       =   "CVG_CONT"
      DataSource      =   "Data1"
      Height          =   315
      Left            =   8670
      TabIndex        =   9
      Tag             =   "40-Coverage Continued"
      Top             =   3540
      Width           =   1125
      _Version        =   65536
      _ExtentX        =   1984
      _ExtentY        =   556
      _StockProps     =   78
      Caption         =   "  "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Value           =   -1  'True
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11355
      _Version        =   65536
      _ExtentX        =   20029
      _ExtentY        =   873
      _StockProps     =   15
      ForeColor       =   0
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
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   22
         Top             =   110
         Width           =   720
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
         Left            =   1320
         TabIndex        =   21
         Top             =   100
         Width           =   1245
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
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
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1005
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmCobra.frx":0000
      Height          =   1995
      Left            =   120
      OleObjectBlob   =   "frmCobra.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   11190
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Coverage Continued ?            "
      Height          =   195
      Index           =   5
      Left            =   6540
      TabIndex        =   28
      Top             =   3630
      Width           =   1530
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "EMPNBR"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9870
      TabIndex        =   27
      Top             =   750
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "COMPNO"
      DataSource      =   "Data1"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   9810
      TabIndex        =   26
      Top             =   1050
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Date Paid To"
      Height          =   195
      Index           =   10
      Left            =   6540
      TabIndex        =   18
      Top             =   4440
      Width           =   945
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "COBRA End Date"
      Height          =   195
      Index           =   9
      Left            =   6540
      TabIndex        =   17
      Top             =   5250
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Payment Due Date"
      Height          =   195
      Index           =   8
      Left            =   6540
      TabIndex        =   16
      Top             =   4830
      Width           =   1350
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Coverage Election Date"
      Height          =   195
      Index           =   7
      Left            =   6540
      TabIndex        =   15
      Top             =   4050
      Width           =   1695
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Reply Date"
      Height          =   195
      Index           =   4
      Left            =   6540
      TabIndex        =   14
      Top             =   3180
      Width           =   795
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Employee Notification Letters"
      Height          =   195
      Index           =   3
      Left            =   330
      TabIndex        =   13
      Top             =   4050
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Employer Notification Date"
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
      Index           =   2
      Left            =   330
      TabIndex        =   12
      Top             =   3600
      Width           =   2280
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Qualifying Reason"
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
      Index           =   1
      Left            =   330
      TabIndex        =   11
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Date of Qualifying Event "
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
      Index           =   0
      Left            =   330
      TabIndex        =   10
      Top             =   3180
      Width           =   2160
   End
End
Attribute VB_Name = "frmCobra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'the most changed in screen jaddy 10/22/99
Option Explicit
Dim LProvs() As Variant
Dim LDepts() As String
Dim LDivs() As Variant
Dim NoDiv As Integer
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim xActin
Dim fglbNew

Function ChkCobra()
Dim x%
ChkCobra = False

If clpCode(1).Caption = "Unassigned" Then
    MsgBox "Invalid Qualifying Reason"
    clpCode(1).SetFocus
    Exit Function
End If
For x% = 1 To 5
    If x% < 3 Then
        If Len(dlpDate(x%).Text) = 0 Then
            MsgBox "Date is Missing"
            dlpDate(x%).SetFocus
            Exit Function
        End If
    End If
    If Len(Trim(dlpDate(x%).Text)) > 0 Then
        If Not IsDate(dlpDate(x%).Text) Then
            MsgBox "Not a valid Date"
            dlpDate(x%).SetFocus
            Exit Function
        End If
    End If
Next x%
'If Len(dlpDate(3)) >= 0 Then
'    If IsDate(dlpDate(2)) Then dlpDate(3) = DateAdd("d", 60, dlpDate(2))
'End If
'If IsDate(dlpDate(4)) Then
'    If Len(dlpDate(6)) = 0 Then dlpDate(6) = DateAdd("d", 45, dlpDate(4))
'    If Len(dlpDate(7)) = 0 Then dlpDate(7) = DateAdd("m", 18, dlpDate(1))
'End If

ChkCobra = True

End Function
Sub ChkCoverCont_Click(Value As Integer)
Dim x%

If Not ChkCoverCont Then
    For x% = 4 To 7
        dlpDate(x%).Text = ""
        dlpDate(x%).Enabled = False
    Next
Else
    dlpDate(4).Enabled = True
    dlpDate(5).Enabled = True
End If

End Sub

Private Sub ChkCoverCont_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err
fglbNew = False

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
xActin = " "

'Call ST_UPD_MODE(False)  ' reset screen's attributes
Call SET_UP_MODE

Me.vbxTrueGrid.SetFocus

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRCOBRA", "Cancel")
Call RollBack

End Sub

Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMCOBRA" Then glbOnTop = ""

End Sub

Public Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim x
xActin = "D"

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "this Record ?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub
    
fglbNew = False

Data1.Recordset.Delete
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
'Call ST_UPD_MODE(False)
Call SET_UP_MODE

Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRCOBRA", "Delete")
Call RollBack

End Sub

Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
Public Sub cmdModify_Click()

On Error GoTo Mod_Err

xActin = "C"

'Call ST_UPD_MODE(True)
Call SET_UP_MODE
'Data1.Recordset.Edit
clpCode(1).Enabled = True
clpCode(1).SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRCOBRA", "Modify")
Call RollBack

End Sub

Sub cmdModify_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdNew_Click()

On Error GoTo AddN_Err

'Call ST_UPD_MODE(True)

fglbNew = True

clpCode(1).SetFocus
xActin = "A"
Data1.Recordset.AddNew
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
lblCNum.Caption = "001"
Call SET_UP_MODE

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

If Err = 3021 Then
    If glbtermopen Then     'Lucy July 5, 2000
        Data1.RecordSource = "Term_HRCOBRA"
    Else
        Data1.RecordSource = "HRCOBRA"
    End If
    fglbEmptyNew = True
    Data1.Refresh
    Data1.Recordset.AddNew
    Resume Next
End If

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRCOBRA", "Add")
Call RollBack

End Sub

Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdOK_Click()

On Error GoTo Add_Err
Dim x As Integer
If Not ChkCobra() Then Exit Sub

If glbtermopen Then Data1.Recordset("TERM_SEQ") = glbTERM_Seq
Call UpdUStats(Me)

'call updDataFld(dlpDate(3), "D", Data1, "REPLY_DATE")
'call updDataFld(dlpDate(4), "D", Data1, "ELECTION_DATE")
'call updDataFld(dlpDate(5), "D", Data1, "DATE_PAID")
'call updDataFld(dlpDate(6), "D", Data1, "PAYMENT_DATE")
'call updDataFld(dlpDate(7), "D", Data1, "END_DATE")
Data1.Recordset("EVENT") = clpCode(1).Text & ""
Data1.Recordset.UpdateBatch
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

fglbNew = False

'Call ST_UPD_MODE(False)
Call SET_UP_MODE
xActin = " "
Me.vbxTrueGrid.SetFocus
If NextFormIF("Cobra Data") Then
    Call cmdNew_Click
End If
Exit Sub

Add_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRCOBRA", "Update")
Call RollBack

End Sub

Sub cmdOK_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Public Sub cmdPrint_Click()
Dim RHeading As String, xReport

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s COBRA Data"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1

End Sub
Public Sub cmdView_Click()
Dim RHeading As String
RHeading = lblEEName & "'s COBRA Data"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading

'Me.vbxCrystal.Password = gstrAccPWord$
'Me.vbxCrystal.UserName = gstrAccUID$
Me.vbxCrystal.Action = 1
End Sub

Sub cmdPrint_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
Sub cmbLetter_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub







Function EERetrieve()
Dim SQLQ As String

On Error GoTo EERError

EERetrieve = False
Screen.MousePointer = HOURGLASS


If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_HRCOBRA"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * from HRCOBRA"
    SQLQ = SQLQ & " where EMPNBR = " & glbLEE_ID
End If

Data1.RecordSource = SQLQ
Data1.Refresh

EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "COBRARetrieve", "HRCOBRA", "SELECT")
Call RollBack
Exit Function

End Function

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber
If ErrorNumber = 3021 Then  ' no record present on a close
    'Response = 0
    ErrorNumber = 0
Else
    Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRDEPNTS", "SELECT")
End If


End Sub

Sub Form_Activate()
    Call SET_UP_MODE
    glbOnTop = "FRMCOBRA"
End Sub

Sub Form_GotFocus()
    glbOnTop = "FRMCOBRA"
End Sub

Sub Form_Load()
Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found

Screen.MousePointer = HOURGLASS
glbOnTop = "FRMCOBRA"

xActin = " "

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If




Screen.MousePointer = DEFAULT

If Not glbtermopen Then
    If glbLEE_ID = 0 Then frmEEFIND.Show 1
    If glbLEE_ID = 0 Then Unload Me: Exit Sub
Else
    If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
    If glbTERM_ID = 0 Then Unload Me: Exit Sub
End If

If EERetrieve() = False Then
    MsgBox "Sorry, Employee can not be found"
    If glbtermopen Then frmTERMEMPL.Show 1 Else frmEEFIND.Show 1
Else
    Me.Show
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
End If
If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "COBRA Maintainance - " & Left$(glbLEE_SName, 8)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = lblEEID

Call cmdModify_Click

Me.vbxTrueGrid.SetFocus

Call INI_Controls(Me)


Screen.MousePointer = DEFAULT

'Call ST_UPD_MODE(False)

End Sub

Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Sub Form_Unload(Cancel As Integer)
    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Set frmCobra = Nothing
    Call NextForm
End Sub


Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If


fUPMode = TF    ' update mode

cmdOK.Enabled = TF
cmdCancel.Enabled = TF

cmdClose.Enabled = FT
cmdModify.Enabled = FT
cmdNew.Enabled = FT
cmdDelete.Enabled = FT
cmdPrint.Enabled = FT
txtLetter(1).Enabled = TF
txtLetter(2).Enabled = TF
txtLetter(3).Enabled = TF
txtLetter(4).Enabled = TF
txtLetter(5).Enabled = TF
clpCode(1).Enabled = TF
dlpDate(1).Enabled = TF
dlpDate(2).Enabled = TF
dlpDate(3).Enabled = TF
dlpDate(4).Enabled = TF
dlpDate(5).Enabled = TF
dlpDate(6).Enabled = TF
dlpDate(7).Enabled = TF
ChkCoverCont.Enabled = TF
'vbxTrueGrid.Enabled = FT
'If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
'   cmdModify.Enabled = False
'   cmdDelete.Enabled = False
'End If
End Sub


Sub dlpDate_LostFocus(Index As Integer)

If Index = 2 Then
  If Len(dlpDate(3)) = 0 Then
    If IsDate(dlpDate(2)) Then dlpDate(3) = DateAdd("d", 60, dlpDate(2))
  End If
End If

If Index = 4 Then
    If IsDate(dlpDate(4)) Then
        If Len(dlpDate(6)) = 0 Then
            dlpDate(6) = DateAdd("d", 45, dlpDate(4))
        End If
        If Len(dlpDate(7)) = 0 Then
            dlpDate(7) = DateAdd("m", 18, dlpDate(1))
        End If
    End If
End If

End Sub

Private Sub txtLetter_GotFocus(Index As Integer)
    Call SetPanHelp(ActiveControl)
End Sub

Sub vbxTrueGrid_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
 Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        If glbtermopen Then         'Lucy July 5, 2000
            SQLQ = "Select * from Term_HRCOBRA"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HRCOBRA"
            SQLQ = SQLQ & " where EMPNBR = " & glbLEE_ID
        End If
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
    If cmdOK.Enabled Then
      cmdOK.SetFocus
    Else
      cmdModify.SetFocus
    End If
End If

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
UpdateRight = True
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
Printable = True
End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    UpdateState = NewRecord
    TF = True
ElseIf Data1.Recordset.EOF And Data1.Recordset.BOF Then
    UpdateState = NoRecord
    TF = False
Else
    UpdateState = OPENING
    TF = True
End If
Call set_Buttons(UpdateState)
If Not UpdateRight Then TF = False
Call ST_UPD_MODE(TF)
End Sub
