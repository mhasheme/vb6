VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmLinamarSkills 
   Caption         =   "Skills for Production"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10095
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtFullCode 
      Appearance      =   0  'Flat
      DataField       =   "SE_FULLCODE"
      Height          =   285
      Left            =   1990
      MaxLength       =   23
      TabIndex        =   0
      Top             =   2880
      Width           =   2235
   End
   Begin VB.TextBox txtCodeDesc 
      Height          =   285
      Left            =   6960
      MaxLength       =   50
      TabIndex        =   5
      Top             =   3780
      Width           =   2055
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SE_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6240
      MaxLength       =   25
      TabIndex        =   13
      Top             =   6660
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SE_LTIME"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4560
      MaxLength       =   25
      TabIndex        =   12
      Top             =   6660
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      Appearance      =   0  'Flat
      DataField       =   "SE_LDATE"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2880
      MaxLength       =   25
      TabIndex        =   11
      Top             =   6660
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox txtSKComment 
      Appearance      =   0  'Flat
      DataField       =   "SE_COMM1"
      Height          =   1545
      Left            =   1995
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Tag             =   "00-Enter comment"
      Top             =   4740
      Width           =   7125
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmLinamarSkills.frx":0000
      Height          =   2145
      Left            =   120
      OleObjectBlob   =   "frmLinamarSkills.frx":0014
      TabIndex        =   9
      Top             =   600
      Width           =   9015
   End
   Begin INFOHR_Controls.DateLookup dlpSKDate 
      DataField       =   "SE_DATE"
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Tag             =   "40-Enter date that skill was acquired"
      Top             =   4380
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   7440
      Top             =   7200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   10
      Top             =   6915
      Width           =   10095
      _Version        =   65536
      _ExtentX        =   17806
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   6495
         Top             =   120
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
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   10095
      _Version        =   65536
      _ExtentX        =   17806
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
         Left            =   6840
         TabIndex        =   30
         Top             =   135
         Width           =   1305
      End
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
         TabIndex        =   17
         Top             =   135
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
         TabIndex        =   16
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEENumber 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   160
         Width           =   1005
      End
   End
   Begin MSMask.MaskEdBox medSKLevel 
      DataField       =   "SE_LEVEL"
      Height          =   285
      Left            =   1995
      TabIndex        =   6
      Tag             =   "10-Enter experience factor"
      Top             =   4080
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
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
   Begin INFOHR_Controls.CodeLookup clpCurrentDIV 
      DataField       =   "SE_CURRENTDIV"
      Height          =   285
      Left            =   7800
      TabIndex        =   1
      Tag             =   "01-Facility - Code"
      Top             =   2880
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpDIV 
      DataField       =   "SE_DIV"
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Tag             =   "01-Facility - Code"
      Top             =   3180
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SE_REGION"
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Tag             =   "01-Facility - Code"
      Top             =   3480
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
      MaxLength       =   8
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "SE_SECTION"
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Tag             =   "01-Facility - Code"
      Top             =   3780
      Width           =   4005
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "HMOP"
      MaxLength       =   12
   End
   Begin VB.Label lblFullCodeDesc 
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   4380
      TabIndex        =   29
      Top             =   2930
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   255
      Left            =   5880
      TabIndex        =   28
      Top             =   3810
      Width           =   975
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Facility:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6600
      TabIndex        =   27
      Top             =   2925
      Width           =   1080
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   1680
      Picture         =   "frmLinamarSkills.frx":4224
      Top             =   2940
      Width           =   240
   End
   Begin VB.Label lblFullCode 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Combined Code"
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
      Left            =   240
      TabIndex        =   26
      Top             =   2940
      Width           =   1335
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Facility"
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
      Left            =   240
      TabIndex        =   25
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label lblCNum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "SE_COMPNO"
      DataSource      =   "Data1"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   330
      TabIndex        =   24
      Top             =   6780
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblEEID 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "SE_EMPNBR"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2160
      TabIndex        =   23
      Top             =   6780
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Product Line"
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
      Left            =   240
      TabIndex        =   22
      Top             =   3540
      Width           =   1095
   End
   Begin VB.Label lblExp 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Experience Factor"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   4140
      Width           =   1290
   End
   Begin VB.Label lblComment 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comment"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   4740
      Width           =   660
   End
   Begin VB.Label lblDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   19
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Operation #"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   18
      Top             =   3840
      Width           =   1305
   End
End
Attribute VB_Name = "frmLinamarSkills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim fUPMode As Integer
Dim fglbNewSalRec%
Dim glbNew
Dim fglbNew
Dim rsDATA As New ADODB.Recordset 'Sam add July 2002 * Remove ADO
Dim Ctrl As Control 'Sam add July 2002 * Remove ADO

Private Function chkSkills()
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, Msg As String

chkSkills = False

On Error GoTo chkSkill_Err

If txtFullCode = "" Then
    MsgBox "Combined code is required field"
    txtFullCode.SetFocus
    Exit Function
End If
If lblFullCodeDesc.Caption = "Unassigned" Then
    If Len(clpDIV.Text) = 0 Or Len(clpCode(1)) = 0 Then
        MsgBox "Combined code must be valid"
        txtFullCode.SetFocus
        Exit Function
    Else
        lblFullCodeDesc.Caption = SaveFullCode
    End If
Else
    lblFullCodeDesc.Caption = SaveFullCode
End If

If Len(medSKLevel) > 0 Then
    If Not IsNumeric(medSKLevel) Then
        MsgBox "Experience Factor is invalid"
        medSKLevel.SetFocus
        Exit Function
    End If

    If medSKLevel > 99 Or medSKLevel < 0 Then
        MsgBox "Experience Factor must be between 0 and 99"
        medSKLevel.SetFocus
        Exit Function
    End If
Else
    medSKLevel = 0
End If

If Len(dlpSKDate.Text) < 1 Then
    dlpSKDate.Text = Format(Now, "short date")    'laura  dec 08, 1997
End If

If Not IsDate(dlpSKDate.Text) Then
    MsgBox "Skill Date is not a valid date"
    dlpSKDate.SetFocus
    Exit Function
End If


chkSkills = True

Exit Function

chkSkill_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkSkill", "HREMPSKL", "edit/Add")
Resume Next

End Function

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err
'data1.Recordset.CancelUpdate
'If Not glbSQL and not glboracle Then Call Pause(0.5)
'data1.Refresh
fglbNew = False
''' Sam add July 2002 * Remove ADO
rsDATA.CancelUpdate
Call Display_Value



'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMPSKL", "Cancel")
Call RollBack '23July99 js

End Sub

'Private Sub cmdCancel_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMESKILLS" Then glbOnTop = ""

End Sub

'Private Sub cmdClose_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String
Dim Skll As String, Skllvl As String, SklDte As String
Dim SQLQ As String, x

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"

a% = MsgBox(Msg, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub


If glbtermopen Then
    gdbAdoIhr001X.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001X.CommitTrans
    Data1.Refresh
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Delete
    gdbAdoIhr001.CommitTrans
    Data1.Refresh
End If
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
    Call Display_Value
End If

fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HREMPSKL", "Delete")
Call RollBack '23July99 js

End Sub




'Private Sub cmdModify_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdNew_Click()
Dim SQLQ As String

fglbNew = True
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
On Error GoTo AddN_Err

Call Set_Control("B", Me)
rsDATA.AddNew

dlpSKDate.Text = Date
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
clpCurrentDIV = Right(lblEEID, 3)
lblCNum.Caption = "001"



glbNew = True
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HREMPSKL", "Add")
Resume Next
End Sub

'Private Sub CmdNew_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x
On Error GoTo Add_Err

If Not chkSkills() Then Exit Sub


Call UpdUStats(Me)
Call Set_Control("U", Me, rsDATA)

If glbtermopen Then
    rsDATA!TERM_SEQ = glbTERM_Seq
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans
Else
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
End If
Data1.Refresh
fglbNew = False
Call SET_UP_MODE
'Call ST_UPD_MODE(False)
If NextFormIF("Skill") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    Data1.Recordset.CancelUpdate
    Data1.Recordset.Resync
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HREMPSKL", "Update")
Resume Next
Unload Me

End Sub

'Private Sub cmdOK_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Skills"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

RHeading = lblEEName & "'s Skills"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub




Private Sub clpCode_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 2
        txtFullCode = clpDIV & "-" & clpCode(1)
        If clpCode(2) <> "" Then txtFullCode = txtFullCode & "-" & clpCode(2)
    End Select
End Sub

Private Sub clpDIV_LostFocus()
    txtFullCode = clpDIV & "-" & clpCode(1)
    If clpCode(2) <> "" Then txtFullCode = txtFullCode & "-" & clpCode(2)
End Sub

Private Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HRSKILLS", "SELECT")

End Sub

Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Screen.MousePointer = HOURGLASS


If glbtermopen Then
    SQLQ = "Select * from LN_Term_EMPSKL"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
Else
    SQLQ = "Select * "
    SQLQ = SQLQ & " from LN_EMPSKL "
    SQLQ = SQLQ & " where SE_EMPNBR = " & glbLEE_ID
End If
SQLQ = SQLQ & " ORDER BY SE_DATE DESC"

Data1.RecordSource = SQLQ
Data1.Refresh


EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "SklsRetrieve", "HREMPSKL", "SELECT")
Resume Next

Exit Function

End Function

Private Sub Form_Activate()
    glbOnTop = "FRMLINAMARSKILLS"
Call SET_UP_MODE
End Sub

Private Sub Form_GotFocus()
    glbOnTop = "FRMLINAMARSKILLS"
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found

glbOnTop = "FRMLINAMARSKILLS"
If glbtermopen Then  'Lucy July 4, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
End If

Screen.MousePointer = HOURGLASS


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




If Len(glbLEE_SName) < 1 Then Exit Sub
Screen.MousePointer = HOURGLASS

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
    Me.Caption = "Skills - " & Left$(glbLEE_SName, 5)
    Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
lblEENum.Caption = ShowEmpnbr(lblEEID)
Call Display_Value
Call SET_UP_MODE
clpCode(1).TextBoxWidth = 800
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

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
Dim Keepfocus As Boolean
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Keepfocus = Not isUpdated(Me)
Cancel = Keepfocus Or (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    MDIMain.panHelp(0).Caption = "Select function from the menu."
    Call NextForm
End Sub

Private Sub lblDesc_Click()

End Sub

Private Sub imgIcon_Click()
txtFullCode_DblClick
End Sub



Private Sub medSKLevel_GotFocus()
    Call SetPanHelp(ActiveControl)
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

fUPMode = TF    ' update mode

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
End If
clpDIV.Enabled = TF
clpCode(1).Enabled = TF
clpCode(2).Enabled = TF
txtCodeDesc.Enabled = TF
clpCurrentDIV.Enabled = TF
txtFullCode.Enabled = TF

medSKLevel.Enabled = TF
dlpSKDate.Enabled = TF
txtSKComment.Enabled = TF
'vbxTrueGrid.Enabled = FT

End Sub


Private Sub medSKLevel_KeyPress(KeyAscii As Integer)
If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
    KeyAscii = 0
    Exit Sub
End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False

End Sub

Private Sub txtFullCode_Change()
Dim rsPROD As New ADODB.Recordset
Dim SQLQ
If txtFullCode = "" Then
    lblFullCodeDesc = ""
    txtCodeDesc = ""
Else
    SQLQ = "SELECT * FROM LN_PROD WHERE TB_FULLCODE='" & txtFullCode & "'"
    rsPROD.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
    
    If rsPROD.EOF Then
        lblFullCodeDesc = "Unassigned"
        txtCodeDesc = ""
    Else
        lblFullCodeDesc = rsPROD("TB_DESC") & ""
        txtCodeDesc = rsPROD("TB_DESC") & ""
    End If
    rsPROD.Close
End If
End Sub

Private Sub txtFullCode_DblClick()
Dim xcodes
glbCode = ""
glbCodeDesc = ""
frmProductLineOperation.Show 1
If Len(glbCode) > 0 Then
    txtFullCode = glbCode
    xcodes = Split(glbCode, "-")
    clpDIV = xcodes(0)
    If Len(clpDIV) = 3 Then
        clpCode(1).TransDiv = clpDIV
        clpCode(2).TransDiv = clpDIV
    End If
    If UBound(xcodes) >= 1 Then clpCode(1) = xcodes(1) Else clpCode(1) = ""
    If UBound(xcodes) >= 2 Then
        Dim c As Long
        Dim tmp As String
        For c = 2 To UBound(xcodes)
            tmp = tmp & xcodes(c) & "-"
        Next c
        clpCode(2) = Left(tmp, Len(tmp) - 1)
    Else
        clpCode(2) = ""
    End If
        
    lblFullCodeDesc = glbCodeDesc
    txtCodeDesc = glbCodeDesc
End If
End Sub

Private Sub txtFullCode_GotFocus()
 Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFullCode_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtFullCode_LostFocus()
Dim xcodes

clpDIV = ""
clpCode(1) = ""
clpCode(2) = ""
If Not Len(txtFullCode) < 1 Then
    xcodes = Split(txtFullCode, "-")
    clpDIV = xcodes(0)
    If Len(clpDIV) = 3 Then
        clpCode(1).TransDiv = clpDIV
        clpCode(2).TransDiv = clpDIV
    End If
    If UBound(xcodes) >= 1 Then
        clpCode(1) = xcodes(1)
        clpCode(1).TransDiv = clpDIV
    End If
    If UBound(xcodes) >= 2 Then
        Dim c As Long
        Dim tmp As String
        For c = 2 To UBound(xcodes)
            tmp = tmp & xcodes(c) & "-"
        Next c
        clpCode(2) = Left(tmp, Len(tmp) - 1)
    Else
        clpCode(2) = ""
    End If
    
End If
End Sub

Private Sub txtSKComment_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtSKDate_Change()
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtSKDate_DblClick()
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Private Sub txtSKDate_GotFocus()
'    Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtSKDate_KeyPress(KeyAscii As Integer)
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

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)

If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub


Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim Skll As String, Skllvl As String, SklDte As String
Dim tdcode$
Dim SQLQ As String

On Error GoTo Tab1_Err
'Sam add july 2002 * remove ado
Call Display_Value

Exit Sub

Tab1_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdrowchng", "HREMPSKL", "Add")
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
''' Sam add July 2002 * Remove Binding Control
Sub Display_Value()
Dim SQLQ
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
        SQLQ = "Select * from LN_Term_EMPSKL"
        SQLQ = SQLQ & " WHERE SE_ID = " & Data1.Recordset!SE_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
    Else
        SQLQ = "Select * "
        SQLQ = SQLQ & " from LN_EMPSKL "
        SQLQ = SQLQ & " where SE_ID = " & Data1.Recordset!SE_ID
        If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
        rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    End If
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End If
Call SET_UP_MODE
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
UpdateRight = gSec_Upd_LinamarSkills
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
ElseIf rsDATA.EOF Then
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

Private Sub lblEEID_Change()

If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
    frmLinamarSkills.Caption = "Skills for Production- " & Left$(glbLEE_SName, 5)
    frmLinamarSkills.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
End If
 If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
'lblEEID = glbLEE_ID
lblEENum = ShowEmpnbr(lblEEID)
If glbLinamar Then  'Ticket #14775
    lblEEProdLine = glbLEE_ProdLine
Else
    lblEEProdLine = ""
End If
End Sub


Private Function SaveFullCode() As String
    Dim rs As New ADODB.Recordset
    Dim strsql As String

    rs.Open "SELECT * FROM LN_PROD WHERE TB_FULLCODE='" & txtFullCode & "'", gdbAdoIhr001, adOpenKeyset, adLockOptimistic, adCmdText
    If rs.EOF Then
        rs.AddNew
        rs("TB_FULLCODE") = txtFullCode.Text
        rs("TB_DIV") = clpDIV.Text
        rs("TB_REGION") = clpCode(1).Text
        rs("TB_SECTION") = clpCode(2).Text
        rs("TB_DESC") = txtCodeDesc.Text
        rs.Update
    Else
        rs("TB_DESC") = txtCodeDesc.Text
        rs.Update
    End If
    rs.Close
    SaveFullCode = txtCodeDesc.Text
End Function



