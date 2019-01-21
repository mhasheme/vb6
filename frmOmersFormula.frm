VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmOmersFormula 
   Caption         =   "OMERS Formula"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   9810
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OM_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   9120
      MaxLength       =   25
      TabIndex        =   25
      Text            =   "LUser"
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OM_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   8640
      MaxLength       =   25
      TabIndex        =   24
      Text            =   "LTime"
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "OM_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   8160
      MaxLength       =   25
      TabIndex        =   23
      Text            =   "Ldate"
      Top             =   3840
      Visible         =   0   'False
      Width           =   495
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   4335
      Width           =   9810
      _Version        =   65536
      _ExtentX        =   17304
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
      Begin VB.CommandButton cmdDelete 
         Appearance      =   0  'Flat
         Caption         =   "&Delete"
         Height          =   375
         Left            =   4350
         TabIndex        =   14
         Tag             =   "Delete Division listed"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdNew 
         Appearance      =   0  'Flat
         Caption         =   "&New"
         Height          =   375
         Left            =   3540
         TabIndex        =   13
         Tag             =   "Create a new Division"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Tag             =   "Cancel changes made"
         Top             =   105
         Width           =   795
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Tag             =   "Save changes made"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdModify 
         Appearance      =   0  'Flat
         Caption         =   "&Edit"
         Height          =   375
         Left            =   960
         TabIndex        =   10
         Tag             =   "Edit the information "
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   135
         TabIndex        =   9
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Appearance      =   0  'Flat
         Caption         =   "&Print"
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Tag             =   "Print Departmental Listing"
         Top             =   105
         Visible         =   0   'False
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6720
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowTitle     =   "Department Codes"
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
         Left            =   7680
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
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
      Bindings        =   "frmOmersFormula.frx":0000
      Height          =   2115
      Left            =   120
      OleObjectBlob   =   "frmOmersFormula.frx":0014
      TabIndex        =   15
      Tag             =   "Division Listings"
      Top             =   120
      Width           =   9585
   End
   Begin MSMask.MaskEdBox medYear 
      DataField       =   "OM_YEAR"
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Tag             =   "10-Year"
      Top             =   2400
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   4
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
   Begin MSMask.MaskEdBox medPPeriodNo 
      DataField       =   "OM_PAYPERIOD_NO"
      Height          =   285
      Left            =   6600
      TabIndex        =   2
      Tag             =   "01-Number of Pay Periods"
      Top             =   2400
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      MaxLength       =   2
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
   Begin MSMask.MaskEdBox medYMPEMax 
      DataField       =   "OM_YMPE_MAX"
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Tag             =   "01-YMPE Maximum"
      Top             =   2760
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPercYMPE 
      DataField       =   "OM_PERCENT_YMPE"
      Height          =   285
      Left            =   6600
      TabIndex        =   4
      Tag             =   "01-Percentage to YMPE"
      Top             =   2760
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medMaxRRP 
      DataField       =   "OM_MAXREG_RRP"
      Height          =   285
      Left            =   2280
      TabIndex        =   5
      Tag             =   "01-Max. Registered RRP"
      Top             =   3120
      Width           =   1530
      _ExtentX        =   2699
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
      Format          =   "$#,##0.00;($#,##0.00)"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medPercRRP 
      DataField       =   "OM_PERC_MAX_RRP"
      Height          =   285
      Left            =   6600
      TabIndex        =   6
      Tag             =   "01-Percentage to Max RRP"
      Top             =   3120
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox medTier3Perc 
      DataField       =   "OM_TIER3_PERC"
      Height          =   285
      Left            =   6600
      TabIndex        =   7
      Tag             =   "01-Tier Three Percentage"
      Top             =   3480
      Width           =   870
      _ExtentX        =   1535
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
      Format          =   "0.00%"
      PromptChar      =   "_"
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tier Three Percentage"
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
      Left            =   4440
      TabIndex        =   22
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage to Max RRP"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   3120
      Width           =   2070
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Max. Registered RRP"
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
      Left            =   120
      TabIndex        =   20
      Top             =   3120
      Width           =   1845
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Percentage to YMPE"
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
      Index           =   3
      Left            =   4440
      TabIndex        =   19
      Top             =   2760
      Width           =   1785
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "YMPE Maximum"
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
      Index           =   2
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Width           =   1350
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Pay Periods"
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
      Index           =   1
      Left            =   4440
      TabIndex        =   17
      Top             =   2400
      Width           =   1950
   End
   Begin VB.Label lblTit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   405
   End
End
Attribute VB_Name = "frmOmersFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbRSOld As String, glbEmptyNew  As Integer
Dim fglbNewRec%
Dim rsDATA As New ADODB.Recordset
Dim Ctrl As Control

Private Sub cmdCancel_Click()
Dim bk
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Set_Control("R", Me, rsDATA)


Call modSTUPD(False)
cmdClose.SetFocus


Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OMERS_FORMULA", "Cancel")
Resume Next

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim Div As String, SQLQ As String, Msg$, a%

On Error GoTo DelErr


Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


gdbAdoIhr001.BeginTrans
rsDATA.Delete
gdbAdoIhr001.CommitTrans
DATA1.Refresh


Exit Sub

DelErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OMERS_FORMULA", "Delete")
Resume Next
End Sub

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call modSTUPD(True)
medYear.SetFocus
fglbNewRec% = False

Exit Sub
Mod_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "Single", "Modify")
Call RollBack

End Sub

Private Sub cmdNew_Click()
glbCodeRef = True

On Error GoTo NewErr

Call modSTUPD(True)

fglbNewRec% = True


Call Set_Control("B", Me)
rsDATA.AddNew

medYear.Enabled = True
medYear.SetFocus

Exit Sub

NewErr:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdNew", "HR_OMERS_FORMULA", "AddNew")
Resume Next

End Sub

Private Sub cmdOK_Click()
Dim xID, ctylist
On Error GoTo OK_Err

If Not chkOmers() Then Exit Sub

Call UpdUStats(Me)

Call Set_Control("U", Me, rsDATA)

gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

xID = rsDATA("OM_ID")

DATA1.Refresh
DATA1.Recordset.Find "OM_ID='" & xID & " '"

fglbNewRec% = False
Call modSTUPD(False)

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OMERS_FORMULA", "Update")
Resume Next
Unload Me

End Sub

Private Function chkOmers()
Dim Div As String, SQLQ As String, Msg$
Dim snapOmers As New ADODB.Recordset
Dim x
chkOmers = False
On Error GoTo chkOmers_Err

If Len(medYear.Text) < 1 Then
    MsgBox ("Year is a required field")
    medYear.SetFocus
    Exit Function
End If
If Len(medPPeriodNo.Text) < 1 Then
    MsgBox ("Number of Pay Periods is a required field")
    medPPeriodNo.SetFocus
    Exit Function
End If

If Len(medYMPEMax.Text) < 1 Then
    MsgBox ("YMPE Maximumis a required field")
    medYMPEMax.SetFocus
    Exit Function
End If
If Len(medPercYMPE.Text) < 1 Then
    MsgBox ("Percentage to YMPE is a required field")
    medPercYMPE.SetFocus
    Exit Function
End If
If Len(medMaxRRP.Text) < 1 Then
    MsgBox ("Max. Registered RRP is a required field")
    medMaxRRP.SetFocus
    Exit Function
End If
If Len(medPercRRP.Text) < 1 Then
    MsgBox ("Percentage to Max RRP is a required field")
    medPercRRP.SetFocus
    Exit Function
End If
If Len(medTier3Perc.Text) < 1 Then
    MsgBox ("Tier Three Percentage is a required field")
    medTier3Perc.SetFocus
    Exit Function
End If

If fglbNewRec% Then
    SQLQ = "SELECT * from HR_OMERS_FORMULA "
    SQLQ = SQLQ & "WHERE OM_YEAR = " & medYear.Text & " "
    
    If snapOmers.State <> 0 Then snapOmers.Close
    snapOmers.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapOmers.BOF And snapOmers.EOF Then
        snapOmers.Close
    Else
        Msg$ = lStr("Duplicate record found!")
        MsgBox Msg$
        snapOmers.Close
        Exit Function
    End If
End If

chkOmers = True

Exit Function

chkOmers_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkOmers", "HR_OMERS_FORMULA", "Cancel")
Resume Next

End Function

Private Sub cmdPrint_Click()
'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

End Sub

Private Sub Form_Load()
Dim SQLQ, I, ctylist, x
glbOnTop = "frmOmersFormula"

DATA1.ConnectionString = glbAdoIHRDB
SQLQ = "SELECT * FROM HR_OMERS_FORMULA "
SQLQ = SQLQ & " ORDER BY OM_YEAR DESC "
DATA1.RecordSource = SQLQ
DATA1.Refresh

Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT
Call modSTUPD(False)
If Not gSec_BenefitGroupSetup Then
    cmdModify.Enabled = False
    cmdNew.Enabled = False
    cmdDelete.Enabled = False
End If

Call Display_Value

End Sub

Private Sub Display_Value()
    Dim SQLQ
    If DATA1.Recordset.EOF Or DATA1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open DATA1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "SELECT * FROM HR_OMERS_FORMULA "
        SQLQ = SQLQ & "WHERE OM_ID='" & DATA1.Recordset!OM_ID & "'"
        SQLQ = SQLQ & " ORDER BY OM_YEAR DESC "
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
        Call Set_Control("R", Me, rsDATA)
    End If
    
End Sub

Private Sub modSTUPD(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

cmdModify.Enabled = FT
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF              '
cmdClose.Enabled = FT

vbxTrueGrid.Enabled = FT
medYear.Enabled = TF
medYMPEMax.Enabled = TF
medMaxRRP.Enabled = TF
medPPeriodNo.Enabled = TF
medPercYMPE.Enabled = TF
medPercRRP.Enabled = TF
medTier3Perc.Enabled = TF
  
End Sub

Private Sub medMaxRRP_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medPercRRP_GotFocus()
Call SetPanHelp(ActiveControl)
If IsNumeric(medPercRRP) Then
    medPercRRP = medPercRRP * 100
End If
End Sub

Private Sub medPercRRP_LostFocus()
If IsNumeric(medPercRRP) Then
    medPercRRP = medPercRRP / 100
End If
End Sub

Private Sub medPercYMPE_GotFocus()
Call SetPanHelp(ActiveControl)
If IsNumeric(medPercYMPE) Then
    medPercYMPE = medPercYMPE * 100
End If
End Sub

Private Sub medPercYMPE_LostFocus()
If IsNumeric(medPercYMPE) Then
    medPercYMPE = medPercYMPE / 100
End If
End Sub

Private Sub medPPeriodNo_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medTier3Perc_GotFocus()
Call SetPanHelp(ActiveControl)
If IsNumeric(medTier3Perc) Then
    medTier3Perc = medTier3Perc * 100
End If
End Sub

Private Sub medTier3Perc_LostFocus()
If IsNumeric(medTier3Perc) Then
    medTier3Perc = medTier3Perc / 100
End If
End Sub

Private Sub medYear_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub medYMPEMax_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub
