VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmLDoors 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "Department Security"
   ClientHeight    =   8310
   ClientLeft      =   480
   ClientTop       =   1050
   ClientWidth     =   11130
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
   ScaleHeight     =   8310
   ScaleWidth      =   11130
   WindowState     =   2  'Maximized
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   4
      Top             =   7650
      Width           =   11130
      _Version        =   65536
      _ExtentX        =   19632
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
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&Update"
         Height          =   375
         Left            =   480
         TabIndex        =   32
         Tag             =   "Print Door Access"
         Top             =   150
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSAdodcLib.Adodc Data1 
         Height          =   405
         Left            =   6960
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   714
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
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowWidth     =   480
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileType   =   2
         BoundReportHeading=   "RGELIST"
         BoundReportFooter=   -1  'True
         WindowControls  =   -1  'True
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "fldoors.frx":0000
      Height          =   2025
      Left            =   300
      OleObjectBlob   =   "fldoors.frx":0014
      TabIndex        =   0
      Top             =   300
      Width           =   10755
   End
   Begin VB.Frame frmDetail 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   60
      TabIndex        =   5
      Top             =   2340
      Width           =   9705
      Begin VB.TextBox txtBadgeID 
         Appearance      =   0  'Flat
         DataField       =   "BADGEID"
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
         Left            =   2690
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "00-Badge ID"
         Top             =   1080
         Width           =   1815
      End
      Begin INFOHR_Controls.CodeLookup clpDIV 
         DataField       =   "DIV"
         Height          =   285
         Left            =   2360
         TabIndex        =   2
         Top             =   750
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "n/a"
         LookupType      =   1
      End
      Begin VB.CheckBox chkEMP 
         Caption         =   "EMP"
         DataField       =   "EMP"
         Height          =   300
         Left            =   6000
         TabIndex        =   30
         Top             =   420
         Visible         =   0   'False
         Width           =   2955
      End
      Begin VB.TextBox txtUSERID 
         Appearance      =   0  'Flat
         DataField       =   "USERID"
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
         Left            =   2670
         TabIndex        =   1
         Tag             =   "10-Enter User ID"
         Top             =   420
         Width           =   1200
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door20"
         DataField       =   "door20"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   19
         Left            =   5580
         TabIndex        =   25
         Top             =   4140
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door19"
         DataField       =   "door19"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   18
         Left            =   5580
         TabIndex        =   24
         Top             =   3840
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door18"
         DataField       =   "door18"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   17
         Left            =   5580
         TabIndex        =   23
         Top             =   3540
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door17"
         DataField       =   "door17"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   16
         Left            =   5580
         TabIndex        =   22
         Top             =   3240
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door16"
         DataField       =   "door16"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   15
         Left            =   5580
         TabIndex        =   21
         Top             =   2940
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door15"
         DataField       =   "door15"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   14
         Left            =   5580
         TabIndex        =   20
         Top             =   2640
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door14"
         DataField       =   "door14"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   13
         Left            =   5580
         TabIndex        =   19
         Top             =   2340
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door13"
         DataField       =   "door13"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   12
         Left            =   5580
         TabIndex        =   18
         Top             =   2040
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door12"
         DataField       =   "door12"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   11
         Left            =   5580
         TabIndex        =   17
         Top             =   1740
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door11"
         DataField       =   "door11"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   10
         Left            =   5580
         TabIndex        =   16
         Top             =   1440
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door10"
         DataField       =   "door10"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   9
         Left            =   360
         TabIndex        =   15
         Top             =   4140
         Width           =   2985
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door9"
         DataField       =   "door9"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   8
         Left            =   360
         TabIndex        =   14
         Top             =   3840
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door8"
         DataField       =   "door8"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   7
         Left            =   360
         TabIndex        =   13
         Top             =   3540
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door7"
         DataField       =   "door7"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   6
         Left            =   360
         TabIndex        =   12
         Top             =   3240
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door6"
         DataField       =   "door6"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   5
         Left            =   360
         TabIndex        =   11
         Top             =   2940
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door5"
         DataField       =   "door5"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   4
         Left            =   360
         TabIndex        =   10
         Top             =   2640
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door4"
         DataField       =   "door4"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   3
         Left            =   360
         TabIndex        =   9
         Top             =   2340
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door3"
         DataField       =   "door3"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   2040
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door2"
         DataField       =   "door2"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   1740
         Width           =   2955
      End
      Begin VB.CheckBox chkDoors 
         Caption         =   "Door1"
         DataField       =   "door1"
         DataSource      =   "Data1"
         Height          =   300
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1440
         Width           =   2955
      End
      Begin VB.CommandButton cmdGrantAll 
         Appearance      =   0  'Flat
         Caption         =   "&Grant All"
         Height          =   360
         Left            =   7200
         TabIndex        =   27
         Top             =   4080
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Badge ID"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   43
         Left            =   360
         TabIndex        =   31
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "User ID or Employee #"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   29
         Top             =   450
         Width           =   1935
      End
      Begin VB.Label lblEEName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Unassigned"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   4140
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label lblTitle 
         Caption         =   "Facility"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   780
         Width           =   1035
      End
   End
   Begin VB.Menu mnu_File 
      Caption         =   "File"
      Begin VB.Menu mnu_Return 
         Caption         =   "&Return to Security"
      End
      Begin VB.Menu mnu_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Printer 
         Caption         =   "Printer Setup"
      End
      Begin VB.Menu mnu_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_File_Exit 
         Caption         =   "&Exit INFO:HR"
      End
   End
   Begin VB.Menu mnu_Find 
      Caption         =   "&Find"
      Visible         =   0   'False
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frmLDoors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ODIV, ODivD, xGlbDiv, xGlbDivDesc
Dim fglbEditMode%
Dim fUPMode As Integer, fglbEmptyNew As Integer
Dim fglbNew
Dim rsDATA As New ADODB.Recordset
Sub cmdCancel_Click()

On Error GoTo Can_Err

Data1.Recordset.CancelUpdate
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh
Call ST_UPD_MODE(False)  ' reset screen's attributes
Me.vbxTrueGrid.SetFocus

fglbNew = False
Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HREMP", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If



End Sub



Sub cmdClose_Click()
Unload Me

End Sub
Sub cmdDelete_Click()
Dim a As Integer, Msg$, INo&, x%

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If

On Error GoTo Del_Err


Msg$ = "Are You Sure You Want To Delete "
Msg$ = Msg$ & Chr(10) & "This Record?  "

a% = MsgBox(Msg$, 36, "Confirm Delete")
If a% <> 6 Then Exit Sub
gdbAdoIhr001.Execute "DELETE FROM LN_DOORS WHERE DIV='" & clpDiv & "' AND USERID ='" & Replace(txtUSERID, "'", "''") & "'"
If Not glbSQL And Not glbOracle Then Call Pause(0.5)
Data1.Refresh

Call ST_UPD_MODE(False)
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OCC_HEALTH_SAFETY", "Delete")
Call RollBack   '10June99 js

End Sub

Private Sub clpDiv_KeyUp(KeyCode As Integer, Shift As Integer)
Call SETUPLABEL
End Sub

Private Sub clpDIV_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Call SETUPLABEL
End Sub

Private Sub cmdGrantAll_Click()
Dim x
For x = 0 To 19
    chkDoors(x) = 1
Next
End Sub

Sub cmdModify_Click()
Dim SQLQ As String


Call ST_UPD_MODE(True)

fglbNew = False
On Error GoTo Edit_Err

txtUSERID.SetFocus

Exit Sub
Edit_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdEdit", "HRJOBEVL", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub



Sub cmdNew_Click()
Dim SQLQ As String
Dim x
fglbNew = True

ST_UPD_MODE (True)
On Error GoTo AddN_Err
txtUSERID = ""
txtBadgeID = ""
clpDiv = ""
For x = 1 To 20
 chkDoors(x - 1) = 0
Next
txtUSERID.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OCC_HEALTH_SAFETY", "Add")
Call RollBack   '10June99 js
End Sub

Sub cmdOK_Click()
Dim x%
Dim xID
On Error GoTo OK_Err

If Not chkDoor Then Exit Sub

If fglbNew Then rsDATA.AddNew
Call Set_Control("U", Me, rsDATA)
gdbAdoIhr001.BeginTrans
rsDATA.Update
gdbAdoIhr001.CommitTrans

Data1.Refresh
fglbNew = False


fglbNew = False
Me.vbxTrueGrid.SetFocus

Exit Sub

OK_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HRJOBEVL", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Sub


Sub cmdPrint_Click()
Dim RHeading As String, xReport, x%

Me.vbxCrystal.WindowTitle = "Door Access Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    Call DoorWRK
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "Rldoors.rpt"
    Me.vbxCrystal.SelectionFormula = "{LN_DOORS.USERID}='" & Replace(txtUSERID, "'", "''") & "' "

    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_DIVISION.Division_Name};ANYCHANGE;A"
Me.vbxCrystal.Destination = 1
Me.vbxCrystal.Action = 1

End Sub

Sub cmdView_Click()
Dim RHeading As String, xReport, x%

'Ticket #24480 - To prevent the change in Printer Setup from info:HR to change the Default Printer
'This setting has been done at the design level but the vbxCrystal.Reset is resetting it so doing it again here.
Me.vbxCrystal.WindowShowPrintSetupBtn = glbCRWPrintSetup

Me.vbxCrystal.WindowTitle = "Door Access Report"
    If glbSQL Or glbOracle Then
        Me.vbxCrystal.Connect = RptODBC_SQL
    End If
    Call DoorWRK
    Me.vbxCrystal.ReportFileName = glbIHRREPORTS & "Rldoors.rpt"
    Me.vbxCrystal.SelectionFormula = "{LN_DOORS.USERID}='" & Replace(txtUSERID, "'", "''") & "' "

    Me.vbxCrystal.GroupCondition(0) = "GROUP1;{@EFullName};ANYCHANGE;A"
    Me.vbxCrystal.GroupCondition(1) = "GROUP2;{HR_DIVISION.Division_Name};ANYCHANGE;A"
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1

End Sub

Private Sub cmdUpdate_Click()
Dim xPath, x9560, x2480, xNoBadgeID
Dim xDiv, xTitle, xPP, i
Dim xDoors, j
xPath = App.Path
xPath = xPath & IIf(Right(xPath, 1) = "/", "", "/")
xNoBadgeID = xPath & "NoBadgeID.txt"
Open xNoBadgeID For Output As #3
Data1.Refresh
If Data1.Recordset.EOF Then Exit Sub
xTitle = "DU(0)=""FD""" & vbNewLine
xTitle = xTitle & "DU(1)=""" & Weekday(Date) & """" & vbNewLine
xTitle = xTitle & "DU(2)=""" & Now & """" & vbNewLine
xTitle = xTitle & "DU(3)=""DN""" & vbNewLine



xDiv = Data1.Recordset!Div
x9560 = xPath & xDiv & "_9560" & ".txt"
x2480 = xPath & xDiv & "_2480" & ".txt"

Open x9560 For Output As #1
Open x2480 For Output As #2
xPP = Replace(Replace(xTitle, "FD", getCtrl("9560", xDiv, 0)), "DN", Data1.Recordset!division_name)
Print #1, xPP
xPP = Replace(Replace(xTitle, "FD", getCtrl("2480", xDiv, 0)), "DN", Data1.Recordset!division_name)
Print #2, xPP
i = 1
Do Until Data1.Recordset.EOF
    
    If xDiv <> Data1.Recordset!Div Then
        Print #1, ""
        Print #2, ""
        Close #1
        Close #2
        xDiv = Data1.Recordset!Div
        x9560 = xPath & xDiv & "_9560" & ".txt"
        x2480 = xPath & xDiv & "_2480" & ".txt"
        
        Open x9560 For Output As #1
        Open x2480 For Output As #2
        xPP = Replace(Replace(xTitle, "FD", getCtrl("9560", xDiv, 0)), "DN", Data1.Recordset!division_name)
        Print #1, xPP
        xPP = Replace(Replace(xTitle, "FD", getCtrl("2480", xDiv, 0)), "DN", Data1.Recordset!division_name)
        Print #2, xPP
        i = 1
    End If
    If IsNull(Data1.Recordset!badgeid) Then
        Print #3, Data1.Recordset!USERID & vbTab & lblEEName
    Else
        xDoors = ","
        For j = 1 To 20
            If getCtrl("9560", xDiv, j) Then
                xDoors = xDoors & "," & IIf(Data1.Recordset("DOOR" & j), j, " ")
            End If
        Next
        Print #1, "DA(" & i & ")=""" & Data1.Recordset!badgeid & vbTab & xDoors & """" & vbNewLine
    
        xDoors = ","
        For j = 1 To 20
            If getCtrl("2480", xDiv, j) Then
                xDoors = xDoors & "," & IIf(Data1.Recordset("DOOR" & j), j, " ")
            End If
        Next
        Print #2, "DA(" & i & ")=""" & Data1.Recordset!badgeid & vbTab & xDoors & """" & vbNewLine
    End If
    Data1.Recordset.MoveNext
    i = i + 1
Loop
Print #1, ""
Print #2, ""
Close #1
Close #2
Close #3
Data1.Refresh
End Sub
Private Function getCtrl(xCtrl, xDiv, xNum)
Dim rsDN As New ADODB.Recordset
Dim SQLQ
Dim i
SQLQ = "select * from LN_DOORS_NAME where DIV='" & xDiv & "'"
rsDN.Open SQLQ, gdbAdoIhr001, adOpenForwardOnly
If xNum <> 0 Then
    If rsDN("DOORCTRL" & xNum) = xCtrl Then
        getCtrl = True
    Else
        getCtrl = False
    End If
Else
    For i = 1 To 20
        If rsDN("DOORCTRL" & i) = xCtrl Then
            getCtrl = i
            Exit Function
        End If
    Next
End If
End Function
Private Sub Form_Activate()
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim Answer, DefVal, Msg, Title  '  variables.
Dim RFound As Integer ' records found
Dim x%
glbOnTop = "FRMLDOORS"
Screen.MousePointer = HOURGLASS


If vbxTrueGrid.Visible Then Me.vbxTrueGrid.SetFocus
Me.Caption = lStr("Doors Access - ")
Call EERetrieve
Call INI_Controls(Me)
Call Display_Value
Screen.MousePointer = DEFAULT


End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select from the menu the appropriate function."
Set frmLDoors = Nothing

End Sub



Public Function EERetrieve()
Dim SQLQ As String

EERetrieve = False

On Error GoTo EERError
Data1.ConnectionString = glbAdoIHRDB
Data1.RecordSource = "select * from LN_DOORS left join hr_division on hr_division.div =ln_doors.div where ln_doors.div in (select div from hr_division where " & glbSeleDiv & ") ORDER BY ln_doors.DIV,USERID"
Data1.Refresh
EERetrieve = True
Exit Function

EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "EERetrieve", "HREMP", "SELECT")
Call RollBack '21June99 js


End Function




Private Sub mnu_File_Exit_Click()
    Call ApplicationEnd
End Sub

'Private Sub mnu_Find_Click()
'Dim SaveID, SaveName, xtxtUSERID, xlblEEName
'    SaveID = glbLUserID
'    SaveName = glbLUserNAME
'    frmUFIND.Show 1
'    If glbEEOK Then
'        xtxtUSERID = glbLUserID
'        xlblEEName = glbLUserNAME
'    Else
'        xtxtUSERID = ""
'        xlblEEName = "Unassigned"
'    End If
'    glbLUserNAME = SaveName
'    glbLUserID = SaveID
'    lblUSERID = xtxtUSERID
'    lblEEName.Caption = xlblEEName
'    glbSecUSERID = glbLUserID
'    glbSecEEName = glbLUserNAME
'End Sub

Private Sub mnu_F_PrintSetup_Click()
MDIMain.vbxCommonDlg.Action = 5

End Sub

Private Sub mnu_Return_Click()
   Unload Me
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

glbOHSEdit% = TF

fUPMode = TF    ' update mode
'cmdOK.Enabled = TF
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdCancel.Enabled = TF
'cmdClose.Enabled = FT
'cmdPrint.Enabled = FT
frmDetail.Enabled = TF
'vbxTrueGrid.Enabled = FT
'If Not gSec_Upd_DoorAccess Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
'End If
End Sub
Private Sub txtUSERID_Change()
Dim rsSR As New ADODB.Recordset
If Len(txtUSERID) > 0 Then
    If chkEMP <> 0 Then
        rsSR.Open "select ED_SURNAME +', '+ED_FNAME as username from HREMP where ED_EMPNBR=" & getEmpnbr(txtUSERID), gdbAdoIhr001, adOpenStatic
    Else
        rsSR.Open "select USERNAME from HR_SECURE_BASIC where USERID='" & Replace(txtUSERID, "'", "''") & "'", gdbAdoIhr001, adOpenStatic
    End If
    If rsSR.EOF Then
        lblEEName.Caption = "Unnasigned"
    Else
        lblEEName.Caption = rsSR("USERNAME")
    End If
    'lblEEName = glbfndEEName(getEmpnbr(txtUSERID), Me)
    lblEEName.Visible = True
Else
    lblEEName.Visible = False
    lblEEName.Caption = "Unnasigned"  'laura dec 06, 1997
End If
End Sub

Private Sub txtUSERID_DblClick()
Dim Msg$
Dim xtxtUSERID, xLblEEName, xTxtEEID
Msg$ = "Select from Employee List?"
If MsgBox(Msg$, vbYesNo, "Select") = vbYes Then
    xTxtEEID = getEmpnbr(txtUSERID)
    xLblEEName = lblEEName.Caption
    frmEEFIND.Show 1
     
    If glbEEOK Then
        xTxtEEID = glbLEE_ID
        xLblEEName = glbLEE_SName & ", " & glbLEE_FName
    Else
        xTxtEEID = ""
        xLblEEName = "Unassigned"
    End If

    
'    Call EmployeeDblClk(xTxtEEID, xLblEEName)
    txtUSERID.Text = ShowEmpnbr(xTxtEEID)
    lblEEName.Caption = xLblEEName
    chkEMP.Value = 1
Else
    Dim SaveID, SaveName
    SaveID = glbLUserID
    SaveName = glbLUserNAME
    frmUFIND.Show 1
    
    If glbEEOK Then
        xtxtUSERID = glbLUserID
        xLblEEName = glbLUserNAME
    Else
        xtxtUSERID = ""
        xLblEEName = "Unassigned"
    End If
    glbLUserNAME = SaveName
    glbLUserID = SaveID
    txtUSERID.Text = xtxtUSERID
    lblEEName.Caption = xLblEEName
    chkEMP.Value = 0
End If
End Sub

Private Sub txtUSERID_GotFocus()
   Call SetPanHelp(Me.ActiveControl)
End Sub


Private Sub txtUSERID_LostFocus()
Dim rsSR As New ADODB.Recordset
If Len(txtUSERID) > 0 Then
    If chkEMP <> 0 Then
        rsSR.Open "select ED_DIV,ED_BADGEID from HREMP where ED_EMPNBR=" & getEmpnbr(txtUSERID), gdbAdoIhr001, adOpenStatic
    Else
        rsSR.Open "select EMPNBR,USERNAME,ED_DIV,ED_BADGEID from HR_SECURE_BASIC left JOIN HREMP ON HR_SECURE_BASIC.EMPNBR=HREMP.ED_EMPNBR where USERID='" & Replace(txtUSERID, "'", "''") & "'", gdbAdoIhr001, adOpenStatic
    End If
    If rsSR.EOF Then
        txtBadgeID = ""
        'txtBadgeID.Enabled = False
    Else
        If IsNull(rsSR("ED_BADGEID")) Then
            txtBadgeID = ""
            txtBadgeID.Enabled = True
        Else
            txtBadgeID = rsSR("ED_BADGEID")
            txtBadgeID.Enabled = False
        End If
    End If
End If
End Sub

Private Sub vbxTrueGrid_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub vbxTrueGrid_HeadClick(ByVal ColIndex As Integer)
Dim SQLQ As String
    
       
        If vbxTrueGrid.Tag = "ASC" Then
            vbxTrueGrid.Tag = "DESC"
        Else
            vbxTrueGrid.Tag = "ASC"
        End If
        
        SQLQ = "select * from LN_DOORS left join hr_division on hr_division.div =ln_doors.div where ln_doors.div in (select div from hr_division where " & glbSeleDiv & ") "
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
End Sub

Private Sub vbxTrueGrid_KeyPress(KeyAscii As Integer)
If KeyAscii = 9 Then ' if the tab key was struck
    KeyAscii = 0
End If

End Sub


Private Sub SETUPLABEL()
Dim rsTD As New ADODB.Recordset
Dim SQLQ, x
rsTD.Open "SELECT * FROM LN_DOORS_NAME WHERE DIV='" & clpDiv & "'", gdbAdoIhr001, adOpenStatic
If rsTD.EOF Then
    For x = 1 To 20
        chkDoors(x - 1).Caption = "Door " & x
        chkDoors(x - 1).Value = 0
        chkDoors(x - 1).Enabled = False
    Next
Else
    For x = 1 To 20
        If IsNull(rsTD("DOORNAME" & x)) Then
            chkDoors(x - 1).Caption = "Door" & x
            chkDoors(x - 1).Value = 0
            chkDoors(x - 1).Enabled = False
        Else
            If Len(rsTD("DOORNAME" & x)) = 0 Then
                chkDoors(x - 1).Caption = "Door" & x
                chkDoors(x - 1).Value = 0
                chkDoors(x - 1).Enabled = False
            Else
                chkDoors(x - 1).Caption = rsTD("DOORNAME" & x)
                chkDoors(x - 1).Enabled = True
                If Not IsNull(rsTD("DOORCTRL" & x)) Then
                    If rsTD("DOORCTRL" & x) = "9560" Then
                        If x Mod 4 = 0 Then
                            chkDoors(x - 1).Enabled = False
                        End If
                    End If
                End If
            End If
        End If
    Next
End If
rsTD.Close
End Sub

Private Function chkDoor()
Dim Div As String, SQLQ As String, Msg$
Dim snapDivs As New ADODB.Recordset

chkDoor = False
On Error GoTo chkDoor_Err

If Len(clpDiv) < 1 Then
    MsgBox lStr("Division Code is a required field")
    clpDiv.SetFocus
    Exit Function
End If


If glbLinamar And (Len(clpDiv) <> 3 Or Not IsNumeric(clpDiv)) Then
    MsgBox lStr("Invalid Division")
    If clpDiv.Enabled Then clpDiv.SetFocus
    Exit Function
End If
If Len(txtUSERID) > 0 Then
    If lblEEName.Caption = "Unnasigned" Then
        MsgBox "If User ID Entered - they must exist"
        txtUSERID.SetFocus
        Exit Function
    End If
Else
    MsgBox lStr("User ID is a required field")
    txtUSERID.SetFocus
    Exit Function
End If
If fglbNew Then
    Div = CStr(clpDiv)
    SQLQ = "SELECT DIV from LN_DOORS "
    SQLQ = SQLQ & "WHERE DIV = '" & Div & "' AND USERID='" & Replace(txtUSERID, "'", "''") & "'"
    
    If snapDivs.State <> 0 Then snapDivs.Close
    snapDivs.Open SQLQ, gdbAdoIhr001, adOpenStatic
    
    If snapDivs.BOF And snapDivs.EOF Then
        snapDivs.Close
    Else
        Msg$ = lStr("This record is duplicate")
        MsgBox Msg$
        snapDivs.Close
        Exit Function
    End If
End If

chkDoor = True

Exit Function

chkDoor_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkDoor", "LN_Door", "Cancel")
Resume Next

End Function
Private Sub DoorWRK()
gdbAdoIhr001.Execute "DELETE FROM LN_DOORS_WRK WHERE WRKEMP='" & glbUserID & "'"
gdbAdoIhr001.Execute "INSERT INTO LN_DOORS_WRK(USEREMP,SHOWNAME,EMP,WRKEMP) SELECT right(ED_EMPNBR,3)+'-'+ left(ED_EMPNBR,LEN(ED_EMPNBR)-3),ED_SURNAME+', '+ED_FNAME,1,'" + glbUserID + "' FROM HREMP"
gdbAdoIhr001.Execute "INSERT INTO LN_DOORS_WRK(USEREMP,SHOWNAME,EMP,WRKEMP) SELECT USERID,USERNAME,0,'" + Replace(glbUserID, "'", "''") + "' FROM HR_SECURE_BASIC"
End Sub

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
        Call SET_UP_MODE
        Exit Sub
    End If
    SQLQ = "select * from LN_DOORS where ID=" & Data1.Recordset!ID
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call SET_UP_MODE
    Call Set_Control("R", Me, rsDATA)
    Call SETUPLABEL
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
UpdateRight = gSec_Upd_DoorAccess
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


Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value
End Sub
