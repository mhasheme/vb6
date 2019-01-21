VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmComPreEmail 
   Caption         =   "Company Preference Email List"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10305
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmailType 
      Appearance      =   0  'Flat
      DataField       =   "VE_TYPE"
      Height          =   285
      Left            =   3720
      MaxLength       =   20
      TabIndex        =   32
      Tag             =   "00-Payroll ID"
      Top             =   2160
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      DataField       =   "VE_EMAIL"
      DataSource      =   " "
      Height          =   795
      Left            =   1740
      MaxLength       =   400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Tag             =   "00-Email Address"
      Top             =   4680
      Width           =   6660
   End
   Begin VB.ComboBox comEmailType 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   1725
      TabIndex        =   0
      Tag             =   "00-Email Type"
      Top             =   2160
      Width           =   1920
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "VE_ADMINBY"
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Tag             =   "00-Administered By"
      Top             =   3450
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDAB"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "VE_ORG"
      Height          =   285
      Index           =   0
      Left            =   6660
      TabIndex        =   10
      Tag             =   "00-Enter Union Code"
      Top             =   3480
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDOR"
   End
   Begin INFOHR_Controls.CodeLookup clpDept 
      DataField       =   "VE_DEPT"
      Height          =   285
      Left            =   1425
      TabIndex        =   2
      Tag             =   "00-Specific Department Desired"
      Top             =   3060
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      MaxLength       =   7
      LookupType      =   2
   End
   Begin INFOHR_Controls.CodeLookup clpDiv 
      DataField       =   "VE_DIV"
      Height          =   285
      Left            =   1425
      TabIndex        =   1
      Tag             =   "00-Specific Division Desired"
      Top             =   2640
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "n/a"
      LookupType      =   1
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "VE_EMP"
      Height          =   285
      Index           =   1
      Left            =   6660
      TabIndex        =   7
      Tag             =   "00-Specific Employment Status Desired"
      Top             =   2610
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDEM"
   End
   Begin INFOHR_Controls.CodeLookup clpPT 
      DataField       =   "VE_PT"
      Height          =   285
      Left            =   6660
      TabIndex        =   8
      Tag             =   "EDPT-Category"
      Top             =   3030
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDPT"
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmComPreEmail.frx":0000
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frmComPreEmail.frx":0014
      TabIndex        =   9
      Top             =   120
      Width           =   9855
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "VE_SECTION"
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Tag             =   "00-Section - Code"
      Top             =   3870
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "VE_LOC"
      Height          =   285
      Index           =   4
      Left            =   6660
      TabIndex        =   11
      Tag             =   "00-Enter Location Code"
      Top             =   3900
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDLC"
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8520
      Top             =   4800
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
      TabIndex        =   25
      Top             =   6075
      Width           =   10305
      _Version        =   65536
      _ExtentX        =   18177
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
         Left            =   4590
         TabIndex        =   31
         Tag             =   "Delete Division listed"
         Top             =   105
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
         Left            =   3780
         TabIndex        =   30
         Tag             =   "Create a new Division"
         Top             =   120
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
         Left            =   2880
         TabIndex        =   29
         Tag             =   "Cancel changes made"
         Top             =   105
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
         Left            =   2040
         TabIndex        =   28
         Tag             =   "Save changes made"
         Top             =   105
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
         Left            =   1200
         TabIndex        =   27
         Tag             =   "Edit the information "
         Top             =   105
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
         Left            =   375
         TabIndex        =   26
         Tag             =   "Close and exit this screen"
         Top             =   105
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   1935
         Top             =   30
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
      End
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "VE_GRPCD"
      Height          =   285
      Index           =   5
      Left            =   6660
      TabIndex        =   6
      Tag             =   "00-Position Group - Code"
      Top             =   2160
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "JBGC"
   End
   Begin Threed.SSCheck chkUserFlag 
      Height          =   255
      Left            =   1740
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   2925
      _Version        =   65536
      _ExtentX        =   5159
      _ExtentY        =   450
      _StockProps     =   78
      Caption         =   "  User Flag"
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
      Font3D          =   3
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      DataField       =   "VE_REGION"
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   5
      Tag             =   "00-Region"
      Top             =   4260
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDRG"
   End
   Begin VB.Label lblRegion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Region"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   34
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Label lblCriteria 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Position Group"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   5160
      TabIndex        =   33
      Top             =   2190
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblTitle 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   24
      Top             =   4725
      Width           =   1095
   End
   Begin VB.Label lblEmailType 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Email Type"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   23
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label lblDiv 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Division"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   22
      Top             =   2640
      Width           =   555
   End
   Begin VB.Label lblDept 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   21
      Top             =   3060
      Width           =   825
   End
   Begin VB.Label lblUnion 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Union"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   20
      Top             =   3510
      Width           =   1020
   End
   Begin VB.Label lblEmpStatys 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employment Status"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   19
      Top             =   2640
      Width           =   1350
   End
   Begin VB.Label lblAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Administered By"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   18
      Top             =   3480
      Width           =   1260
   End
   Begin VB.Label lblSelCri 
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
      Left            =   120
      TabIndex        =   17
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label lblPT 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5160
      TabIndex        =   16
      Top             =   3060
      Width           =   630
   End
   Begin VB.Label lblSection 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Section"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   150
      TabIndex        =   15
      Top             =   3900
      Width           =   1260
   End
   Begin VB.Label lblLocation 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5190
      TabIndex        =   14
      Top             =   3930
      Width           =   615
   End
End
Attribute VB_Name = "frmComPreEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew As Boolean
Dim fglbSDate As Variant
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim RSDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control
Dim UpdateState As UpdateStateEnum
Dim fglbESQLQ As String
Dim fglbVSQLQ As String

Private Sub cmdCancel_Click()
Dim bk

On Error GoTo Can_Err

fglbNew = False

RSDATA.CancelUpdate

Call Display_Value
Call ST_UPD_MODE(False)

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPREEMAIL", "Cancel")
Resume Next

End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
Dim a As Integer, Msg As String

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err

Msg = "Are You Sure You Want To Delete "
Msg = Msg & "This Record?"
a% = MsgBox(Msg, 36, "Confirm Delete")

If a% <> 6 Then Exit Sub

gdbAdoIhr001.BeginTrans
RSDATA.Delete
gdbAdoIhr001.CommitTrans
Data1.Refresh


'Call SET_UP_MODE
Call ST_UPD_MODE(False)


Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HRPREEMAIL", "Delete")
Call RollBack '09June99 js

End Sub

Private Sub cmdModify_Click()
On Error GoTo Mod_Err

Call ST_UPD_MODE(True)

clpDiv.SetFocus

Exit Sub

Mod_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "HRPREEMAIL", "Modify")
Call RollBack

End Sub

Private Sub cmdNew_Click()
On Error GoTo AddN_Err

Call Set_Control("B", Me)
RSDATA.AddNew

fglbNew = True
'Call SET_UP_MODE
Call ST_UPD_MODE(True)

If Len(glbEmalType) > 0 Then
    txtEmailType.Text = glbEmalType
End If
clpDiv.SetFocus
Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HRPREEMAIL", "Add")
Call RollBack '09June99 js
End Sub

Private Sub cmdOK_Click()
Dim X%
Dim bmk As Variant

On Error GoTo cmdOK_Err
If (Data1.Recordset.EOF And Data1.Recordset.BOF) Then
    bmk = 0
Else
    bmk = Data1.Recordset.Bookmark
End If

If Not chkEPayroll() Then Exit Sub

Call Set_Control("U", Me, RSDATA)

fglbNew = False

gdbAdoIhr001.BeginTrans
RSDATA.Update
gdbAdoIhr001.CommitTrans
Data1.Refresh
If Not bmk = 0 Then
    Data1.Recordset.Bookmark = bmk
End If

Call Display_Value

Call ST_UPD_MODE(False)

Me.vbxTrueGrid.Enabled = True
Me.vbxTrueGrid.SetFocus
Screen.MousePointer = DEFAULT

Exit Sub

cmdOK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRPREEMAIL", "Update")
Call RollBack '09June99 js

End Sub

Private Sub comEmailType_LostFocus()
txtEmailType.Text = comEmailType.Text
End Sub

Private Sub Form_Activate()
'Call SET_UP_MODE
'Me.cmdModify_Click
End Sub

Private Sub Form_Load()
MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False

Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim I%, SQLQ

'Me.Show
glbOnTop = "frmComPreEmail"

Call combType

Screen.MousePointer = HOURGLASS

If glbSamuel Then 'Ticket #23453 Franks 03/25/2013
    If glbEmalType = "Termination" Then
        chkUserFlag.Caption = " Send email to Rept. Authority"
        chkUserFlag.DataField = "VE_USER_FLAG"
        chkUserFlag.Visible = True
    End If
End If

Data1.ConnectionString = glbAdoIHRDB
    
SQLQ = "SELECT * FROM HRPREEMAIL "
If Len(glbEmalType) > 0 Then
    SQLQ = SQLQ & "WHERE VE_TYPE = '" & glbEmalType & "' "
    Me.Caption = "Company Preference Email List: " & lStr(glbEmalType)
End If
SQLQ = SQLQ & "ORDER BY VE_TYPE,VE_ADMINBY,VE_DIV,VE_DEPT,VE_SECTION  "
Data1.RecordSource = SQLQ

Data1.Refresh

Call setRptCaption(Me)

Screen.MousePointer = DEFAULT

Call Display_Value
Call ST_UPD_MODE(False)
                                               '
vbxTrueGrid.Columns(1).Caption = lStr("Division")
vbxTrueGrid.Columns(2).Caption = lStr("Department")
vbxTrueGrid.Columns(3).Caption = lStr("Administered By")
vbxTrueGrid.Columns(4).Caption = lStr("Section")
vbxTrueGrid.Columns(5).Caption = lStr("Region")
vbxTrueGrid.Columns(7).Caption = lStr("Category")
vbxTrueGrid.Columns(8).Caption = lStr("Union")
vbxTrueGrid.Columns(9).Caption = lStr("Location")
Call INI_Controls(Me)

If glbEmalType = "Position" Then
    lblCriteria(5).Visible = True
    clpCode(5).Visible = True
End If

If Len(glbEmalType) > 0 Then
    txtEmailType.Text = glbEmalType
End If

Screen.MousePointer = DEFAULT                           '
End Sub

Private Sub combType()
comEmailType.AddItem "New Hire"
comEmailType.AddItem "Position" 'Ticket #21444 Franks 02/09/2012
comEmailType.AddItem "Salary"
comEmailType.AddItem "Benefits"
comEmailType.AddItem "Termination"
comEmailType.AddItem "Rehire"
comEmailType.AddItem "Leave Changes"
comEmailType.AddItem "Performance"
comEmailType.AddItem "Dependent"
comEmailType.AddItem "ESS-Request Approval"
comEmailType.AddItem "Address Changes"  ''8.0 - Ticket #22682 - Employee Address Change
comEmailType.AddItem "New Applicant"    ''8.0 - Ticket #25273 - New Applicant - ATS
comEmailType.AddItem "ESS-Request Submit"   'Ticket #27060 - S.U.C.C.E.S.S.
comEmailType.AddItem "Employee Flags"   ''Ticket #26934 - Oshawa Community Health Centre - Employee Flags
comEmailType.AddItem "H&S Incident"   ''Ticket #28664 Franks 05/30/2016
End Sub

Private Sub Display_Value()
    Dim SQLQ
    
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
        RSDATA.Open Data1.RecordSource, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        'Call SET_UP_MODE
        Exit Sub
    End If
    
    SQLQ = "SELECT * FROM HRPREEMAIL where VE_ID= " & Data1.Recordset!VE_ID
    If RSDATA.State <> 0 Then: If RSDATA.EOF Then RSDATA.Close Else If RSDATA.EditMode = adEditAdd Then RSDATA.CancelUpdate: RSDATA.Close Else RSDATA.Close
    RSDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    
    If RSDATA.EOF Or RSDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, RSDATA)
    'Call SET_UP_MODE

End Sub

'Public Property Get ChangeAction() As UpdateStateEnum
'If fGLBNew Then
'    ChangeAction = NewRecord
'Else
'    ChangeAction = OPENING
'End If
'End Property
'Public Property Let ChangeAction(vData As UpdateStateEnum)
'If vData = NewRecord Then fGLBNew = True
'End Property
'
'Public Property Get RelateMode() As RelateModeEnum
'RelateMode = RelateSetUp
'End Property
'
'Public Property Get UpdateRight() As Boolean
'UpdateRight = True
'End Property
'
'Public Property Get Addable() As Boolean
'Addable = True
'End Property
'Public Property Get Updateble() As Boolean
'Updateble = True
'End Property
'Public Property Get Deleteble() As Boolean
'Deleteble = True
'End Property
'Public Property Get Printable() As Boolean
'Printable = False 'True
'End Property

Public Sub SET_UP_MODE()
Dim TF As Boolean
Dim UpdateState As UpdateStateEnum
If fglbNew Then
    'UpdateState = NewRecord
    TF = True
ElseIf Data1.Recordset.EOF Then
    'UpdateState = NoRecord
    TF = False
Else
    'UpdateState = OPENING
    TF = True
End If
Call ST_UPD_MODE(TF)
'Call set_Buttons(UpdateState)
'If Not UpdateRight Then TF = False
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

fUPMode = TF

cmdModify.Enabled = FT           '
cmdDelete.Enabled = FT          '
cmdNew.Enabled = FT             '
cmdClose.Enabled = FT
cmdCancel.Enabled = TF          '
cmdOK.Enabled = TF
If Len(glbEmalType) > 0 Then
    comEmailType.Enabled = False
Else
    comEmailType.Enabled = TF            '
End If
clpDiv.Enabled = TF
clpDept.Enabled = TF
clpPT.Enabled = TF
txtEmail.Enabled = TF
clpCode(0).Enabled = TF
clpCode(1).Enabled = TF      '
clpCode(2).Enabled = TF      '
clpCode(3).Enabled = TF      '
clpCode(4).Enabled = TF
clpCode(5).Enabled = TF
clpCode(6).Enabled = TF
If glbSamuel Then 'Ticket #23453 Franks 03/25/2013
    chkUserFlag.Enabled = TF
End If

If Data1.Recordset.BOF Or Data1.Recordset.EOF Then
    cmdModify.Enabled = False
    cmdDelete.Enabled = False
End If

End Sub

Private Function RollBack()
On Error GoTo RR
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
RR:
End Function

Private Sub txtEmailType_Change()
    Select Case txtEmailType
        Case "": comEmailType.ListIndex = -1
        Case "New Hire":
            comEmailType.ListIndex = 0
        Case "Position":
            comEmailType.ListIndex = 1 'Ticket #21444 Franks 02/09/2012
        Case "Salary":
            comEmailType.ListIndex = 2 '1
        Case "Benefits":
            comEmailType.ListIndex = 3 '2
        Case "Termination":
            comEmailType.ListIndex = 4 '3
        Case "Rehire":
            comEmailType.ListIndex = 5 '4
        Case "Leave Changes":
            comEmailType.ListIndex = 6 '5
        Case "Performance":
            comEmailType.ListIndex = 7 '6
        Case "Dependent":
            comEmailType.ListIndex = 8 '7
        Case "ESS-Request Approval":
            comEmailType.ListIndex = 9 '8
        Case "Address Changes":
            comEmailType.ListIndex = 10 '9  '8.0 - Ticket #22682 - Employee Address Change
        Case "New Applicant"
            comEmailType.ListIndex = 11 '10  '8.0 - Ticket #25273 - New Applicant - ATS
        Case "ESS-Request Submit":        'Ticket #27060 - S.U.C.C.E.S.S.
            comEmailType.ListIndex = 12 '11
        Case "Employee Flags"
            comEmailType.ListIndex = 13 '12   'Ticket #26934 - Oshawa Community Health Centre - Employee Flags
        Case "H&S Incident"
            comEmailType.ListIndex = 14 'Ticket #28664 Franks 05/30/2016
    End Select

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

If Data1.Recordset.EOF Or Data1.Recordset.BOF = 0 Then
    Exit Sub
End If
End Sub

Private Function chkEPayroll()
Dim rsVT As New ADODB.Recordset
Dim Msg As String
Dim SQLQ As String
Dim X%, xchk

chkEPayroll = False

If Len(comEmailType.Text) = 0 Then
    MsgBox ("Email Type is required field.")
    'comEmailType.SetFocus
    Exit Function
End If

For X% = 0 To 4
If Len(clpCode(X%).Text) > 0 And clpCode(X%).Caption = "Unassigned" Then
    MsgBox "If Code entered it must be valid."
    clpCode(X%).SetFocus
    Exit Function
End If
Next X%

If Len(clpDept.Text) > 0 And clpDept.Caption = "Unassigned" Then
    MsgBox lStr("If Department Entered - it must be valid.")
     clpDept.SetFocus
    Exit Function
End If
If Len(clpDiv.Text) < 1 Then
    'If glbDIVCount = 1 And glbLinamar Then
    '    MsgBox lStr("Division is required field")
    '     clpDiv.SetFocus
    '    Exit Function
    'End If
Else
    If clpDiv.Caption = "Unassigned" Then
        MsgBox lStr("If Division Entered - it must be valid.")
         clpDiv.SetFocus
        Exit Function
    End If
End If

If clpPT.Caption = "Unassigned" Then
    MsgBox "If " & lblPT.Caption & " Entered - it must be valid."
    clpPT.SetFocus
    Exit Function
End If

If Len(txtEmail.Text) = 0 Then
    MsgBox "Email Address is required."
    txtEmail.SetFocus
    Exit Function
End If


Call getWSQLQ   '"C")
SQLQ = "SELECT * FROM HRPREEMAIL WHERE " & fglbVSQLQ
rsVT.Open SQLQ, gdbAdoIhr001, adOpenStatic
If Not rsVT.EOF Then
    MsgBox "You can not add duplicate record"
    clpDiv.SetFocus
    Exit Function
End If
    

chkEPayroll = True

Exit Function

chkMUEntitle_Err:

End Function

Private Sub getWSQLQ() 'xType)
Dim xDiv, xDept, xORG, xAsOf, xEMP, xEmpMode, xGRPCE
Dim xLoc, xSection
Dim xFromDate
Dim xToDate
Dim xID
Dim SQLQ As String
fglbESQLQ = "" 'glbSeleDeptUn
fglbVSQLQ = " (1=1) "
If Len(clpDiv.Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & "AND (VE_DIV IS NULL OR VE_DIV='') "
Else
    fglbVSQLQ = fglbVSQLQ & "AND VE_DIV = '" & clpDiv.Text & "' "
End If
If Len(clpDept.Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_DEPT IS NULL OR VE_DEPT='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_DEPT = '" & clpDept.Text & "' "
End If
If Len(clpCode(0).Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_ORG IS NULL OR VE_ORG='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_ORG = '" & clpCode(0).Text & "' "
End If

If Len(clpCode(1).Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_EMP IS NULL OR VE_EMP='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_EMP = '" & clpCode(1).Text & "' "
End If
If Len(clpPT.Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_PT IS NULL OR VE_PT='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_PT = '" & clpPT.Text & "' "
End If

If Len(clpCode(2).Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_ADMINBY IS NULL OR VE_ADMINBY='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_ADMINBY = '" & clpCode(2).Text & "' "
End If
If Len(clpCode(3).Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_SECTION IS NULL OR VE_SECTION='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_SECTION = '" & clpCode(3).Text & "' "
End If
If Len(clpCode(4).Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_LOC IS NULL OR VE_LOC='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_LOC = '" & clpCode(4).Text & "' "
End If
'Ticket #21444 Franks 02/09/2012
''If Len(xGRPCE) = 0 Then
''    fglbVSQLQ = fglbVSQLQ & " AND (VE_GRPCD IS NULL OR VE_GRPCD='') "
''Else
''    fglbVSQLQ = fglbVSQLQ & " AND VE_GRPCD = '" & xGRPCE & "' "
''End If
If Len(clpCode(5).Text) = 0 Then 'Position Group
    fglbVSQLQ = fglbVSQLQ & " AND (VE_GRPCD IS NULL OR VE_GRPCD='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_GRPCD = '" & clpCode(5).Text & "' "
End If
'Ticket #27515 Franks 09/06/2015
If Len(clpCode(6).Text) = 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND (VE_REGION IS NULL OR VE_REGION='') "
Else
    fglbVSQLQ = fglbVSQLQ & " AND VE_REGION = '" & clpCode(6).Text & "' "
End If
If fglbNew Then
    xID = 0
Else
    If Not RSDATA.EOF Then
        xID = RSDATA("VE_ID")
    Else
        xID = 0
    End If
End If
If xID > 0 Then
    fglbVSQLQ = fglbVSQLQ & " AND NOT VE_ID = " & xID & " "
End If
If Len(glbEmalType) > 0 Then
    fglbVSQLQ = fglbVSQLQ & "AND VE_TYPE = '" & glbEmalType & "' "
End If
'getWSQLQ = fglbVSQLQ

End Sub


