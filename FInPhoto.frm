VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form FRMINPHOTO 
   AutoRedraw      =   -1  'True
   Caption         =   "Maintain Photos"
   ClientHeight    =   8730
   ClientLeft      =   15
   ClientTop       =   1020
   ClientWidth     =   13590
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   13590
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optExpDelPhoto 
      Caption         =   "Export / Delete Photo from info:HR database"
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   600
      Width           =   3615
   End
   Begin VB.OptionButton optImportPhoto 
      Caption         =   "Import Photo into info:HR database"
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   120
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.Frame frImpPhoto 
      Caption         =   "Import Photo"
      Height          =   1335
      Left            =   360
      TabIndex        =   21
      Top             =   1080
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CheckBox chkFile 
         Caption         =   "File Names are equal to Employee Numbers"
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   3555
      End
      Begin VB.CheckBox chkReplace 
         Caption         =   "Replace Existing Photo"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Value           =   1  'Checked
         Width           =   2355
      End
   End
   Begin VB.CheckBox chkImpWord 
      Caption         =   "Import Resume File"
      Height          =   315
      Left            =   6240
      TabIndex        =   5
      Top             =   1320
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.Frame frmDelExpPhotos 
      Caption         =   "Export/Delete Photos"
      Height          =   1815
      Left            =   360
      TabIndex        =   15
      Top             =   1080
      Width           =   5775
      Begin VB.CheckBox chkExportPhotos 
         Caption         =   "Export ALL existing Employees Photos out of info:HR database"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1310
         Width           =   4875
      End
      Begin VB.CheckBox chkDeleteAll 
         Caption         =   "DELETE ALL EMPLOYEES PHOTOS from info:HR database"
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
         Left            =   240
         TabIndex        =   1
         Top             =   795
         Width           =   5475
      End
      Begin VB.CheckBox chkDelete 
         Caption         =   "Delete Existing Photo from info:HR database"
         Height          =   375
         Left            =   240
         TabIndex        =   0
         Top             =   280
         Width           =   3675
      End
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   3465
      Left            =   2340
      TabIndex        =   11
      Top             =   4260
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2340
      TabIndex        =   10
      Top             =   3930
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.FileListBox File1 
      Height          =   3795
      Left            =   5280
      MultiSelect     =   2  'Extended
      Pattern         =   "*.jpg"
      TabIndex        =   12
      Top             =   3930
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   480
      Left            =   0
      TabIndex        =   13
      Top             =   8250
      Width           =   13590
      _Version        =   65536
      _ExtentX        =   23971
      _ExtentY        =   847
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
         Left            =   8520
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
      End
   End
   Begin VB.Frame frmFile 
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   300
      TabIndex        =   16
      Top             =   4500
      Visible         =   0   'False
      Width           =   7695
      Begin INFOHR_Controls.EmployeeLookup elpEEID 
         Height          =   315
         Left            =   1740
         TabIndex        =   8
         Top             =   90
         Width           =   5000
         _ExtentX        =   8811
         _ExtentY        =   556
         RefreshDescriptionWhen=   2
      End
      Begin VB.TextBox txtFileName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         MaxLength       =   32
         TabIndex        =   9
         Tag             =   "00-File Name (Do not Enter Extension TXT)"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblEENum 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Number"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   150
         Width           =   1290
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Do Not Enter the file Extension (Must be 'JPG')."
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   3720
         TabIndex        =   18
         Top             =   480
         Width           =   4260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Import File Name"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   525
         Width           =   1620
      End
   End
   Begin INFOHR_Controls.EmployeeLookup elpEEID_Del 
      Height          =   285
      Left            =   2040
      TabIndex        =   6
      Tag             =   "10-Enter Employee Number"
      Top             =   3240
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   503
      ShowUnassigned  =   1
      TextBoxWidth    =   6355
      RefreshDescriptionWhen=   2
      MultiSelect     =   -1  'True
      Enabled         =   0   'False
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   7
      Tag             =   "00-Section"
      Top             =   3580
      Visible         =   0   'False
      Width           =   4200
      _ExtentX        =   7408
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "EDSE"
   End
   Begin VB.Label lblSec 
      AutoSize        =   -1  'True
      Caption         =   "Section"
      Height          =   195
      Left            =   420
      TabIndex        =   25
      Top             =   3625
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Note: Do not 'Export/Delete Photo' out of info:HR database if you are using the ESS/Timesheet web modules)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3960
      TabIndex        =   24
      Top             =   690
      Visible         =   0   'False
      Width           =   9450
   End
   Begin VB.Label lblEENum 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Number"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   420
      TabIndex        =   20
      Top             =   3285
      Width           =   1290
   End
   Begin VB.Label lblPath 
      BackStyle       =   0  'Transparent
      Caption         =   "Export to Path"
      Height          =   195
      Left            =   420
      TabIndex        =   14
      Top             =   3990
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "FRMINPHOTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FPath, UPDTCNT
Dim ImportFile As String
Dim xDeleteable

Private Sub chkDelete_Click()
    If chkDelete.Value = vbChecked Then
        chkFile.Value = 0
        chkReplace.Value = 0
        lblPath.Visible = False
        Drive1.Visible = False
        Dir1.Visible = False
        File1.Visible = False
        frmFile.Visible = False
        
        lblEENum(1).Visible = True
        elpEEID_Del.Visible = True
        
        '8.0 - Ticket #22682 - Remove photos from HR_PHOTO to independent folder.
        chkDeleteAll.Value = vbUnchecked
        chkDeleteAll.Enabled = False
        chkExportPhotos.Value = vbUnchecked
        chkExportPhotos.Enabled = False
                
        elpEEID_Del.Enabled = True
        lblEENum(1).Enabled = True
        
    '8.0 - Ticket #22682 - Remove photos from HR_PHOTO to independent folder. Don't need option to replace/import photos
    Else
        '8.0 - Ticket #22682 - Remove photos from HR_PHOTO to independent folder.
        chkDeleteAll.Enabled = True
        chkExportPhotos.Enabled = True
                
        elpEEID_Del.Text = ""
        elpEEID_Del.Enabled = False
        lblEENum(1).Enabled = False
    End If
    Call SET_UP_MODE
End Sub

Private Sub chkDeleteAll_Click()
    If chkDeleteAll.Value = vbChecked Then
        'Clear the employee #s entered to delete the photos of because this is DELETE ALL EMPLOYEES photo.
        elpEEID_Del.Text = ""
        
        elpEEID_Del.Enabled = False
        lblEENum(1).Enabled = False
        
        chkDelete.Value = vbUnchecked
        chkDelete.Enabled = False
        chkExportPhotos.Value = vbUnchecked
        chkExportPhotos.Enabled = False
        
    ElseIf chkDelete.Value = vbChecked Then
        elpEEID_Del.Enabled = True
        lblEENum(1).Enabled = True
        
        chkExportPhotos.Value = vbUnchecked
        chkExportPhotos.Enabled = False
    Else
        chkDelete.Enabled = True
        chkExportPhotos.Enabled = True
    End If
    Call SET_UP_MODE
End Sub

Private Sub chkExportPhotos_Click()
    If chkExportPhotos.Value = vbChecked Then
        chkDelete.Value = vbUnchecked
        chkDeleteAll.Value = vbUnchecked
        
        chkDelete.Value = vbUnchecked
        chkDeleteAll.Value = vbUnchecked
        
        chkDelete.Enabled = False
        chkDeleteAll.Enabled = False
        
        'Allow folder to export selection
        lblPath.Visible = True
        Drive1.Visible = True
        Dir1.Visible = True
        
        'Ticket #26315 Franks 11/26/2014 - Jerry asked to make this function generic in 8.1
        'If glbWFC Then 'Ticket #26308 Franks 11/21/2014
            Call WFCPhotoScreen(True)
        'End If
    Else
        chkDelete.Enabled = True
        chkDeleteAll.Enabled = True
        
        lblPath.Visible = False
        Drive1.Visible = False
        Dir1.Visible = False
        
        'Ticket #26315 Franks 11/26/2014 - Jerry asked to make this function generic in 8.1
        'If glbWFC Then 'Ticket #26308 Franks 11/21/2014
            Call WFCPhotoScreen(False)
        'End If
    End If
    
    Call SET_UP_MODE
End Sub

Private Sub WFCPhotoScreen(xFlag) 'Ticket #26308 Franks 11/21/2014
    elpEEID_Del.Enabled = xFlag 'True
    lblEENum(1).Enabled = xFlag 'True
    lblSec.Visible = xFlag
    clpCode(0).Visible = xFlag
    lblSec.Caption = lStr("Section")
End Sub

'Private Sub optExpDelPhoto_Click()
''Ticket #22682 - Disable the Import Photo option if Employee Photo In Other Folder is checked (Company Preference - File Locations)
'If optExpDelPhoto Then
'    frmDelExpPhotos.Visible = True
'    frmDelExpPhotos.Enabled = True
'    frImpPhoto.Visible = False
'Else
'    If gsEMPLOYEEPHOTO Then
'        optImportPhoto.Enabled = False
'
'        optExpDelPhoto.Value = vbChecked
'        frmDelExpPhotos.Visible = True
'        frmDelExpPhotos.Enabled = True
'        frImpPhoto.Visible = False
'    Else
'        frmDelExpPhotos.Visible = False
'        optImportPhoto.Enabled = True
'        frImpPhoto.Visible = True
'    End If
'End If
'
'End Sub

Private Sub chkFile_Click()
    frmFile.Visible = 1 - chkFile.Value
    xDeleteable = 1 - chkFile.Value
    
    Call SET_UP_MODE
    
    'cmdDelete.Visible = 1 - chkFile.Value
    If chkFile Then
        chkDelete.Value = vbUnchecked
    End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdClose_GotFocus()
 Call SetPanHelp(Me.ActiveControl)
End Sub

Public Sub cmdDelete_Click()
Dim SQLQ As String, X
Dim Title$, Msg$, DgDef As Variant, Response%
Dim xESQLQ As String

Title = "Employee Photo Delete"


On Error GoTo Mod_Err
If chkDelete.Value = vbUnchecked And chkDeleteAll.Value = vbUnchecked Then
    MsgBox "To delete employee's Photo in info:HR database, 'Delete Existing Photo from info:HR database' or 'DELETE ALL EMPLOYEES PHOTOS from info:HR database' should be checked.", vbExclamation, "Delete Employee Photo from info:HR database"
    Exit Sub
End If
If chkDelete.Value = vbChecked Then    '8.0 - Ticket #22682
    If Len(elpEEID_Del.Text) = 0 Then
        MsgBox "Employee Number is required."
        elpEEID_Del.SetFocus
        Exit Sub
    End If
End If
If Not elpEEID_Del.ListChecker Then
    Exit Sub
End If

'If elpEEID.Caption = "Unassinged" Then
'    MsgBox "Employee Number is not valid."
'    elpEEID.SetFocus
'    Exit Sub
'End If
    
DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
'Msg$ = "Are you sure you want to Delete " & elpEEID.Caption & "'s Photo?"
'8.0 - Ticket #22682 - Remove photos from HR_PHOTO table because we are moving to folder now
If chkDeleteAll.Value = vbChecked Then
    Msg$ = "Are you sure you want to DELETE ALL Employees Photos from info:HR database?"
Else
    Msg$ = "Are you sure you want to Delete Employee's Photo from info:HR database?"
End If
Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
If Response% = IDNO Then Exit Sub

If chkDeleteAll.Value = vbChecked Then
    MDIMain.panHelp(0).Caption = "Deleting Employees Photos from info:HR database, please wait....."
Else
    MDIMain.panHelp(0).Caption = "Deleting Employee's Photo from info:HR database, please wait....."
End If

Screen.MousePointer = HOURGLASS

'8.0 - Ticket #22682 - Remove photos from HR_PHOTO table because we are moving to folder now
If chkDeleteAll.Value = vbChecked Then
    'As per Department security
    xESQLQ = glbSeleDeptUn
    gdbAdoIhr001.Execute "DELETE FROM HR_PHOTO WHERE PT_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & xESQLQ & ")"
Else
    gdbAdoIhr001.Execute "DELETE FROM HR_PHOTO WHERE PT_EMPNBR IN (" & getEmpnbr(elpEEID_Del) & ")"
End If

Screen.MousePointer = DEFAULT

MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

'8.0 - Ticket #22682 - Remove photos from HR_PHOTO table because we are moving to folder now
If chkDeleteAll.Value = vbChecked Then
    MsgBox "ALL Employees photos DELETED from info:HR database successfully."
Else
    MsgBox "Employee's photo Deleted from info:HR database successfully."
End If

Exit Sub

Mod_Err:
If Err = 53 Then Resume Next

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDelete", "Photo", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Sub cmdModify_Click()
Dim SQLQ As String, X
Dim Title$, Msg$, DgDef As Variant, Response%

If chkImpWord Then
    Title = "Employee Resume Import"
    'If Not gSec_Import_Attendance Then
    '    MsgBox "You Do Not Have Authority For This Transacaction"
    '    Exit Sub
    'End If
    
    On Error GoTo Mod_Err
    
    If Not chkPhoto() Then Exit Sub
    
    DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
    Msg$ = "Are you sure you want to Import Resume?"
    Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
    If Response% = IDNO Then Exit Sub
    
    Screen.MousePointer = HOURGLASS
    
    ChDir FPath
    If Not modUpdateSelectionResume() Then GoTo bpMod
    
    MDIMain.panHelp(0).FloodPercent = 100
    
    Close
    '-----------------------------------------------------
    
    Screen.MousePointer = DEFAULT
    MDIMain.panHelp(0).FloodType = 0
    MDIMain.panHelp(1).Caption = " Update Completed"
    MDIMain.panHelp(2).Caption = ""
    If UPDTCNT = 0 Then
        Msg$ = "No Photo Imported "
    Else
        Msg$ = Str(UPDTCNT)
        If UPDTCNT = 1 Then Msg$ = Msg$ & " Record " Else Msg$ = Msg$ & " Records "
        Msg$ = Msg$ & "Imported Successfully "
    End If
    DgDef = MB_ICONINFORMATION
    MsgBox Msg$, DgDef, Title
Else
    If optImportPhoto Then
        Title = "Employee Photo Import"
        'If Not gSec_Import_Attendance Then
        '    MsgBox "You Do Not Have Authority For This Transacaction"
        '    Exit Sub
        'End If
    
        On Error GoTo Mod_Err
        
        If Not chkPhoto() Then Exit Sub
        
        DgDef = MB_YESNO + MB_ICONSTOP + MB_DEFBUTTON2  ' Describe dialog.
        Msg$ = "Are you sure you want to Import Photos into info:HR database?"
        Response% = MsgBox(Msg$, DgDef, Title)    ' Get user response.
        If Response% = IDNO Then Exit Sub
    
        Screen.MousePointer = HOURGLASS
    
        ChDir FPath
        If Not modUpdateSelection() Then GoTo bpMod
    
        MDIMain.panHelp(0).FloodPercent = 100
    
        Close
        '-----------------------------------------------------
    
        Screen.MousePointer = DEFAULT
        MDIMain.panHelp(0).FloodType = 0
        MDIMain.panHelp(1).Caption = " Update Completed"
        MDIMain.panHelp(2).Caption = ""
        If UPDTCNT = 0 Then
            Msg$ = "No Photo Imported "
        Else
            Msg$ = Str(UPDTCNT)
            If UPDTCNT = 1 Then Msg$ = Msg$ & " Record " Else Msg$ = Msg$ & " Records "
            Msg$ = Msg$ & "Imported Successfully "
        End If
        DgDef = MB_ICONINFORMATION
        MsgBox Msg$, DgDef, Title
    Else
        'Export Photos to a file
        Call Export_Photos
    End If
End If

bpMod:

Screen.MousePointer = DEFAULT
MDIMain.panHelp(0).FloodType = 0
MDIMain.panHelp(1).Caption = ""
MDIMain.panHelp(2).Caption = ""

Exit Sub

Mod_Err:
If Err = 53 Then Resume Next

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdmod", "USTI Import", "Update")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Private Sub Export_Photos()
    Dim xPath As String
    Dim Response%
    
    'Make sure Export existing photos is selected
    If chkExportPhotos.Value = vbChecked Then
        'Get the user to enter the path to export the Photos to and
        'Update that path to Company Pref. 'EMPLOYEEPHOTOPATH'
        xPath = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\"))
    
        'Verify the export folder
        'Ticket #26315 Franks 11/26/2014 - Jerry asked to make this function generic in 8.1
        'If glbWFC Then 'Ticket #26308 Franks 11/21/2014
            Response% = MsgBox("Are you sure you want to export these Employees Photos from info:HR database to '" & xPath & "' folder?", vbQuestion + vbYesNo, "Confirm Employees Photos export folder")
        'Else
        '    Response% = MsgBox("Are you sure you want to export all Employees Photos from info:HR database to '" & xPath & "' folder?", vbQuestion + vbYesNo, "Confirm Employees Photos export folder")
        'End If
        If Response% = vbNo Then Exit Sub
        
        'Export photos
        Screen.MousePointer = HOURGLASS
        
        MDIMain.panHelp(0).Caption = "Exporting Photos, please wait...."
        
        Call Export_Photos_FromDB(xPath)
        
        MDIMain.panHelp(0).Caption = ""
    
        Screen.MousePointer = DEFAULT
    Else
        MsgBox "To 'Export/Delete Photo from info:HR database', one of the 'Export/Delete Photo' checkboxes should be checked.", vbExclamation
    End If
End Sub

Private Sub Export_Photos_FromDB(xAppPath)
    Dim AppPath
    Dim rsPhoto As New ADODB.Recordset
    Dim byteChunk() As Byte
    
    Dim FileNumber As Integer
    Dim TempFile As String
    Dim TempDir As String * 255
            
    Dim rsPrefer As New ADODB.Recordset
    Dim SQLQ As String
    Dim xESQLQ As String
    
    'Path user selected to export the Photos into
    AppPath = xAppPath
    
    'As per Department security
    xESQLQ = glbSeleDeptUn
    
    'Retrieve Photos of each employee
    SQLQ = "SELECT * FROM HR_PHOTO WHERE PT_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE " & xESQLQ & ")"
    'Ticket #26315 Franks 11/26/2014 - Jerry asked to make this function generic in 8.1
    'If glbWFC Then 'Ticket #26308 Franks 11/21/2014
        If Len(elpEEID_Del.Text) > 0 Then
            SQLQ = SQLQ & " AND PT_EMPNBR IN (" & getEmpnbr(elpEEID_Del) & ") "
        End If
        If clpCode(0).Visible And Len(clpCode(0).Text) > 0 Then
            SQLQ = SQLQ & " AND PT_EMPNBR IN (SELECT ED_EMPNBR FROM HREMP WHERE ED_SECTION = '" & clpCode(0).Text & "') "
        End If
    'End If
    rsPhoto.Open SQLQ, gdbAdoIhr001, adOpenStatic, adLockOptimistic
    If rsPhoto.EOF Then Exit Sub
    Do While Not rsPhoto.EOF
        'Set File Name using the Employee #
        TempFile = AppPath & rsPhoto("PT_EMPNBR") & ".jpg"
        
        'If file already exists, delete it
        If (Dir(TempFile)) <> "" Then Kill TempFile
        
        FileNumber = FreeFile
        Open TempFile For Binary Access Write As FileNumber
        
        ReDim byteChunk(rsPhoto("PT_PHOTO").ActualSize)
        byteChunk() = rsPhoto("PT_PHOTO").GetChunk(rsPhoto("PT_PHOTO").ActualSize)
        Put FileNumber, , byteChunk()
    
        Close FileNumber
        
        rsPhoto.MoveNext
    Loop
    rsPhoto.Close
    Set rsPhoto = Nothing
        
    'Update Company Pref with Employee Photo path
    If glbWFC Then 'Ticket #26308 Franks 11/21/2014
        Screen.MousePointer = DEFAULT
        MsgBox "   Finished!   "
    Else
        SQLQ = "SELECT * FROM HRPREFERENCE WHERE HP_FUN_NAME = 'EMPLOYEEPHOTOPATH'"
        rsPrefer.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
        If Not rsPrefer.EOF Then
            rsPrefer("HP_EMAIL") = AppPath
            rsPrefer.Update
        End If
        rsPrefer.Close
        Set rsPrefer = Nothing
        
        MDIMain.panHelp(0).Caption = "Employees Photos export completed."
        
        Screen.MousePointer = DEFAULT
        
        MsgBox "Employees Photos exported from info:HR database successfully." & vbCrLf & vbCrLf & "To view Employee's Photo in info:HR, please turn it ON from the 'Company Preference' screen under the 'Setup' menu in info:HR.", vbInformation, "Turn-ON Employee Photo view option"
    End If
End Sub

Private Sub optExpDelPhoto_Click()
    If optExpDelPhoto Then
        frImpPhoto.Visible = False
        frmDelExpPhotos.Visible = True
        
        chkDelete.Value = vbUnchecked
        chkDeleteAll.Value = vbUnchecked
        chkExportPhotos.Value = vbUnchecked
        
        chkDelete.Enabled = True
        chkDeleteAll.Enabled = True
        chkExportPhotos.Enabled = True
        
        'Allow folder to export selection
        lblPath.Caption = "Export to Path"
        File1.Pattern = "*.jpg"
        lblPath.Visible = False
        Drive1.Visible = False
        Dir1.Visible = False
        File1.Visible = False
                
        frmFile.Visible = False
                
        lblEENum(1).Visible = True
        elpEEID_Del.Visible = True
        
        'Feb 11th 2014: From the Meeting with Jerry and Mostafa, since Mostafa said ESS/Timesheet cannot read
        'a folder thatdoes not reside on the web server which needs to use Network Service account (local account)
        'to access any local folders from web module, we decided to put this message saying clients with web
        'modules should not move the photos out of info:HR database.
        lblMsg.Visible = True
        
    Else
        'Feb 11th 2014: From the Meeting with Jerry and Mostafa, since Mostafa said ESS/Timesheet cannot read
        'a folder thatdoes not reside on the web server which needs to use Network Service account (local account)
        'to access any local folders from web module, we decided to put this message saying clients with web
        'modules should not move the photos out of info:HR database.
        lblMsg.Visible = False
        
'        If gsEMPLOYEEPHOTO Then
'            optImportPhoto.Enabled = False
'
'            optExpDelPhoto.Value = vbChecked
'            frmDelExpPhotos.Visible = True
'            frmDelExpPhotos.Enabled = True
'            frImpPhoto.Visible = False
'
'            frmFile.Visible = False
'            lblPath.Visible = False
'
'            lblPath.Caption = "Export to Path"
'            Drive1.Visible = False
'            Dir1.Visible = False
'            File1.Visible = False
'
'            lblEENum(1).Visible = True
'            elpEEID_Del.Visible = True
'        Else
            frImpPhoto.Visible = True
            frmDelExpPhotos.Visible = False
            
            frmFile.Top = 2940
            frmFile.Visible = False
            lblPath.Caption = "Import From Path"
            
            lblPath.Visible = True
            Drive1.Visible = True
            Dir1.Visible = True
            File1.Visible = True
            
            lblEENum(1).Visible = False
            elpEEID_Del.Visible = False
            
            'optImportPhoto.Enabled = True
'        End If
    End If
    Call SET_UP_MODE
End Sub

Private Sub optImportPhoto_Click()
    If optImportPhoto Then
        frImpPhoto.Visible = True
        frmDelExpPhotos.Visible = False
        
        frmFile.Top = 2940
        frmFile.Visible = False
        lblPath.Caption = "Import From Path"
        
        lblPath.Visible = True
        Drive1.Visible = True
        Dir1.Visible = True
        File1.Visible = True
                
        lblEENum(1).Visible = False
        elpEEID_Del.Visible = False
        
        'Feb 11th 2014: From the Meeting with Jerry and Mostafa, since Mostafa said ESS/Timesheet cannot read
        'a folder thatdoes not reside on the web server which needs to use Network Service account (local account)
        'to access any local folders from web module, we decided to put this message saying clients with web
        'modules should not move the photos out of info:HR database.
        lblMsg.Visible = False
        
    Else
        'Feb 11th 2014: From the Meeting with Jerry and Mostafa, since Mostafa said ESS/Timesheet cannot read
        'a folder thatdoes not reside on the web server which needs to use Network Service account (local account)
        'to access any local folders from web module, we decided to put this message saying clients with web
        'modules should not move the photos out of info:HR database.
        lblMsg.Visible = True
        
        frImpPhoto.Visible = False
        frmDelExpPhotos.Visible = True
        
        chkDelete.Value = vbUnchecked
        chkDeleteAll.Value = vbUnchecked
        chkExportPhotos.Value = vbUnchecked
        
        chkDelete.Enabled = True
        chkDeleteAll.Enabled = True
        chkExportPhotos.Enabled = True
        
        lblPath.Caption = "Export to Path"
        lblPath.Visible = False
        File1.Pattern = "*.jpg"
        Drive1.Visible = False
        Dir1.Visible = False
        File1.Visible = False
                
        frmFile.Visible = False
        
        lblEENum(1).Visible = True
        elpEEID_Del.Visible = True
    End If
    Call SET_UP_MODE
End Sub

Private Sub chkImpWord_Click()
If chkImpWord Then
    optImportPhoto.Enabled = False
    optExpDelPhoto.Enabled = False
    
    File1.Pattern = "*.*"
    
    chkDelete.Value = False
    chkDeleteAll.Value = False
    chkExportPhotos.Value = False
    
    chkDelete.Visible = False
    chkDeleteAll.Visible = False
    chkExportPhotos.Visible = False
    
    frImpPhoto.Caption = "Import Resume"
    chkReplace.Caption = "Replace Existing Resume"
    chkReplace.Enabled = False
    
    frmFile.Visible = True
    txtFileName.Text = ""
    txtFileName.Enabled = False
    Label2.Visible = False
    frmFile.Top = 2940 ' lblEENum(1).Top
    lblPath.Visible = True
    lblPath.Caption = "Import From Path"
    Drive1.Visible = True
    Dir1.Visible = True
    File1.Visible = True
         
    lblEENum(1).Visible = False
    elpEEID_Del.Visible = False
Else
    optImportPhoto.Enabled = True
    optExpDelPhoto.Enabled = True
    frImpPhoto.Caption = "Import Photo"
    chkReplace.Enabled = True
    
    If gsEMPLOYEEPHOTO Then
        optImportPhoto.Enabled = False
        optExpDelPhoto.Value = vbChecked
        
        Call optExpDelPhoto_Click
    Else
        optImportPhoto.Enabled = True
        frImpPhoto.Visible = True
        frmDelExpPhotos.Visible = False
        
        Call optImportPhoto_Click
    End If
    
    File1.Pattern = "*.jpg"
    
    'chkDelete.Visible = True
    'chkDeleteAll.Visible = True
    'chkExportPhotos.Visible = True
    'chkReplace.Value = False
    
    'frmFile.Visible = False
    'lblPath.Visible = False
    'lblPath.Caption = "Export to Path"
    'Drive1.Visible = False
    'Dir1.Visible = False
    'File1.Visible = False
    
    'lblEENum(1).Visible = True
    'elpEEID_Del.Visible = True

End If

End Sub

Private Sub chkReplace_Click()
    If chkReplace Then
        chkDelete.Value = vbUnchecked
    End If
End Sub

Private Sub Dir1_Change()
    ChDir Dir1.Path
    File1.Path = Dir1.Path
    File1.Pattern = "*.JPG"
End Sub

Private Sub Drive1_Change()
Dim xdir, xerror
On Error GoTo CKERROR
xerror = False
Dir1.Path = Drive1.Drive
Exit Sub
CKERROR:
    If Err = 68 Then
         MsgBox "Invalid Drive Selected"
         Drive1.Drive = App.Path
         xerror = True
         Resume Next
    End If
    MsgBox "ERROR " & Str(Err)
    xerror = True
    Resume Next

End Sub

Private Sub File1_Click()
    Dim iit As Integer
    Dim ii1 As Long
    Dim sit As String
    For iit = 0 To File1.ListCount - 1
        If File1.selected(iit) Then
            sit = File1.List(iit)
            If chkImpWord Then
                txtFileName.Text = UCase(File1.List(iit))
            Else
                ii1 = InStr(sit, ".")
                If ii1 > 0 Then
                    sit = Mid(sit, 1, ii1 - 1)
                    txtFileName.Text = UCase(sit)
                Else
                    txtFileName.Text = UCase(File1.List(iit))
                End If
            End If
        End If
    Next
    
    ' dkostka - 10/16/2001 - Shouldn't be able to select multiple files if you are choosing
    '   pictures one by one.  Can't change this at runtime via control property so have to
    '   do it in code.
    If chkFile.Value = 0 Then
        For ii1 = 0 To File1.ListCount - 1
            If ii1 <> File1.ListIndex Then File1.selected(ii1) = False
        Next ii1
    End If
End Sub


Private Sub Form_Activate()
Call INI_Controls(Me)
Call SET_UP_MODE
End Sub

Private Sub Form_Load()
Dim X%
Dim Y%

On Error GoTo Line_Err

glbOnTop = "FRMINPHOTO"
Screen.MousePointer = HOURGLASS
  'sets eenames Laura Oct 30, 1997
Screen.MousePointer = DEFAULT

'Drive1.Drive = "c:"
'Dir1.Path = "c:\"

Drive1.Drive = "G:"
Dir1.Path = "G:\"
FPath = Dir1.Path
'Ticket #25676 Franks 07/08/2014 - remove it
'If glbWFC Then 'For test
'    chkImpWord.Visible = True
'End If

'Ticket #22682 - Disable the Import Photo option if Employee Photo In Other Folder is checked (Company Preference - File Locations)
If gsEMPLOYEEPHOTO Then
    optImportPhoto.Enabled = False
    optExpDelPhoto.Value = vbChecked
    
    Call optExpDelPhoto_Click
Else
    optImportPhoto.Enabled = True
    frImpPhoto.Visible = True
    frmDelExpPhotos.Visible = False
    
    Call optImportPhoto_Click
End If

Exit Sub

Line_Err:
    If Err = "68" Then
        'MsgBox Err.Description
        Resume Next
    End If
    
End Sub

Private Sub Form_LostFocus()
MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIMain.panHelp(0).Caption = "Select function from the menu."
End Sub


Sub txtFileName_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub txtFileName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Function modUpdateSelectionResume()
Dim xxx, xx1, X%, XCNT
Dim xEMPNBR, xShowEmpNbr
Dim SQLQ
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%, SSERIAL
Dim rsEmp As New ADODB.Recordset
Dim xPath, xFileName As String

On Error GoTo modUpdateSelection_Err

modUpdateSelectionResume = False

UPDTCNT = 0
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(0).FloodType = 1

MDIMain.panHelp(0).FloodPercent = 0

If False Then
    Call AppendPhoto(getEmpnbr(elpEEID), ImportFile)
Else
    xPath = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\"))
    For X = 0 To File1.ListCount - 1
        If X <> 0 Then
            MDIMain.panHelp(0).FloodPercent = (X / (File1.ListCount - 1)) * 100
        ElseIf (X = 0) And (File1.ListCount - 1) = 0 Then
            MDIMain.panHelp(0).FloodPercent = 100
        End If

        If File1.selected(X) Then
            xFileName = UCase(File1.List(X))
            'xShowEmpNbr = elpEEID 'xFileName 'Left(xFileName, InStr(xFileName, ".JPG") - 1)
            xShowEmpNbr = Left(xFileName, InStr(xFileName, ".") - 1)
            xEMPNBR = getEmpnbr(xShowEmpNbr)
            If Not IsNumeric(xEMPNBR) Then xEMPNBR = 0
            If xEMPNBR <> 0 Then
                rsEmp.Open "SELECT ED_EMPNBR FROM HREMP where ED_EMPNBR=" & xEMPNBR & " AND " & glbSeleDeptUn, gdbAdoIhr001, adOpenStatic
                If Not rsEmp.EOF Then
                    xFileName = xPath & xFileName
                    Call AppendResume(xEMPNBR, xFileName, Right(xFileName, 3))
                    File1.selected(X) = False
                End If
                rsEmp.Close
            End If
        End If
        DoEvents
    Next
End If

MDIMain.panHelp(0).Caption = ""
modUpdateSelectionResume = True
Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update", "ImportPhoto", "Import")
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then Resume Next Else Unload Me
End Function

Function modUpdateSelection()
Dim xxx, xx1, X%, XCNT
Dim xEMPNBR, xShowEmpNbr
Dim SQLQ
Dim Msg$, Title$, DgDef As Variant
Dim Response%, pct%, SSERIAL
Dim rsEmp As New ADODB.Recordset
Dim xPath, xFileName As String
On Error GoTo modUpdateSelection_Err
modUpdateSelection = False

UPDTCNT = 0
MDIMain.panHelp(2).Caption = ""
MDIMain.panHelp(1).Caption = " Please Wait"
MDIMain.panHelp(0).FloodType = 1

MDIMain.panHelp(0).FloodPercent = 0

If chkFile = 0 Then
    Call AppendPhoto(getEmpnbr(elpEEID), ImportFile)
Else
    xPath = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\"))
    For X = 0 To File1.ListCount - 1
        If X <> 0 Then
            MDIMain.panHelp(0).FloodPercent = (X / (File1.ListCount - 1)) * 100
        ElseIf (X = 0) And (File1.ListCount - 1) = 0 Then
            MDIMain.panHelp(0).FloodPercent = 100
        End If

        If File1.selected(X) Then
            xFileName = UCase(File1.List(X))
            xShowEmpNbr = Left(xFileName, InStr(xFileName, ".JPG") - 1)
            xEMPNBR = getEmpnbr(xShowEmpNbr)
            If Not IsNumeric(xEMPNBR) Then xEMPNBR = 0
            If xEMPNBR <> 0 Then
                rsEmp.Open "SELECT ED_EMPNBR FROM HREMP where ED_EMPNBR=" & xEMPNBR & " AND " & glbSeleDeptUn, gdbAdoIhr001, adOpenStatic
                If Not rsEmp.EOF Then
                    xFileName = xPath & xFileName
                    Call AppendPhoto(xEMPNBR, xFileName)
                    File1.selected(X) = False
                End If
                rsEmp.Close
            End If
        End If
        DoEvents
    Next
End If

MDIMain.panHelp(0).Caption = ""
modUpdateSelection = True
Screen.MousePointer = DEFAULT

Exit Function

modUpdateSelection_Err:

Screen.MousePointer = DEFAULT
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Update", "ImportPhoto", "Import")
Screen.MousePointer = DEFAULT

If gintRollBack% = False Then Resume Next Else Unload Me
End Function

Function chkPhoto()
Dim Alphabet, xlen, I%, xwk, xok
chkPhoto = False
On Error GoTo chkPhoto_Err

If chkFile = 0 Then
    If chkDelete Then
        MsgBox "Mass Update cannot be done when 'Delete Existing Photo' is checked. Click Mass Delete button to delete the Photo."
        elpEEID_Del.SetFocus
        Exit Function
    End If
    
    If optImportPhoto Then
        If chkFile = 0 And chkReplace = 0 Then
            MsgBox "To 'Import Photo into info:HR database', one of the 'Import Photo' checkboxes must be checked."
            optImportPhoto.SetFocus
            Exit Function
        End If
    End If
    
    If Len(txtFileName) = 0 Then
        MsgBox "File Name is required."
        File1.SetFocus
        Exit Function
    End If
    
    txtFileName = LTrim(txtFileName)
    xlen = Len(txtFileName)
    If chkImpWord.Value = vbUnchecked Or chkImpWord.Visible = False Then
        ' dkostka - 10/16/2001 - Added space and -_()! to end of alphabet, filenames can have these chars
        Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890-_()! "
        xok = True
        For I% = 1 To xlen
            xwk = Mid(txtFileName, I%, 1)
            If InStr(Alphabet, xwk) = 0 Then
                xok = False
                Exit For
            End If
        Next
        If Not xok Then
            MsgBox "Invalid File Name"
            'txtFileName.SetFocus
            File1.SetFocus
            Exit Function
        End If
    End If
    
    ' dkostka - 10/16/2001 - A valid employee number is required.
    If Len(elpEEID.Text) = 0 Then
        MsgBox "Employee Number is required."
        elpEEID.SetFocus
        Exit Function
    End If
    If elpEEID.Caption = "Unassinged" Then
        MsgBox "Employee Number is not valid."
        elpEEID.SetFocus
        Exit Function
    End If
    ImportFile = UCase(Dir1.Path) & UCase(IIf(Right(Dir1.Path, 1) = "\", "", "\")) & UCase(txtFileName & ".JPG")
    'MsgBox ImportFile
    If Dir(ImportFile) = "" Then
        MsgBox "FILE not Found :" & Chr(10) & "[" & ImportFile & "]"
        txtFileName.SetFocus
        Exit Function
    End If
    
End If
chkPhoto = True

Exit Function

chkPhoto_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkPhoto", "HR_Photo", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Function

Public Sub AppendResume(zEMPNBR, FileName As String, FileExtension As String)
    Dim rsPhoto As New ADODB.Recordset

    Dim byteChunk() As Byte
    Dim X, xChr
    Dim FileNumber As Integer
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    rsPhoto.Open "select * from HRDOC_EMP WHERE RE_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsPhoto.EOF Then
        If chkReplace = 0 Then
            Exit Sub
        Else
            rsPhoto.Delete
        End If
    End If
    UPDTCNT = UPDTCNT + 1
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))
   
    rsPhoto.AddNew
    rsPhoto("RE_EMPNBR") = zEMPNBR
    rsPhoto("RE_COMPNO") = "001"
    rsPhoto("RE_FILEEXT") = FileExtension
    rsPhoto("RE_TYPE") = "RESUME"
    rsPhoto("RE_LUSER") = glbUserID
    rsPhoto("RE_LDATE") = Date
    rsPhoto("RE_LTIME") = Time$
    Get FileNumber, , byteChunk
    rsPhoto!RE_DOC.AppendChunk byteChunk
    Close FileNumber
    
    If glbSQL Or glbOracle Then rsPhoto.Update
'    rsPHOTO.Requery
    rsPhoto.Close

End Sub

Public Sub AppendPhoto(zEMPNBR, FileName As String)

    Dim rsPhoto As New ADODB.Recordset

    Dim byteChunk() As Byte
    Dim X, xChr
    Dim FileNumber As Integer
    If Not IsNumeric(zEMPNBR) Then Exit Sub
    rsPhoto.Open "select * from HR_PHOTO WHERE PT_EMPNBR=" & zEMPNBR, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
    If Not rsPhoto.EOF Then
        If chkReplace = 0 Then
            Exit Sub
        Else
            Do
                rsPhoto.Delete
                rsPhoto.MoveNext
            Loop Until rsPhoto.EOF
        End If
    End If
    UPDTCNT = UPDTCNT + 1
    FileNumber = FreeFile
    Open FileName For Binary Access Read As FileNumber
    ReDim byteChunk(FileLen(FileName))
   
    rsPhoto.AddNew
    rsPhoto("PT_EMPNBR") = zEMPNBR
    rsPhoto("PT_COMPNO") = "001"
    rsPhoto("PT_LUSER") = glbUserID
    rsPhoto("PT_LDATE") = Date
    rsPhoto("PT_LTIME") = Time$
    Get FileNumber, , byteChunk
    rsPhoto!PT_PHOTO.AppendChunk byteChunk
    Close FileNumber
    If glbSQL Or glbOracle Then rsPhoto.Update
'    rsPHOTO.Requery
    rsPhoto.Close

End Sub

Public Property Get ChangeAction() As UpdateStateEnum
ChangeAction = OPENING
End Property

Public Property Get RelateMode() As RelateModeEnum
RelateMode = MassChanges
End Property

Public Property Get UpdateRight() As Boolean
UpdateRight = True
End Property

Public Property Get Addable() As Boolean
Addable = False
End Property

Public Property Get Updateble() As Boolean
If (optImportPhoto.Enabled And optImportPhoto) Or (chkExportPhotos) Or (optExpDelPhoto And Not chkDelete And Not chkDeleteAll) Then
    Updateble = True
ElseIf chkDelete Or chkDeleteAll Then
    Updateble = False
End If
If chkDelete Or chkDeleteAll Then
    Updateble = False
End If
'Updateble = (optImportPhoto.Enabled And optImportPhoto) Or chkExportPhotos Or Not chkDelete Or Not chkDeleteAll
End Property

Public Property Get Deleteble() As Boolean
If chkExportPhotos Or (optImportPhoto.Enabled And optImportPhoto) Then
    Deleteble = False
Else
    Deleteble = True
End If
End Property

Public Property Get Printable() As Boolean
Printable = False
End Property

Public Sub SET_UP_MODE()
Call set_Buttons
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If glbUserUploadMode = UploadFormWithoutCheck And UnloadMode = 1 Then Exit Sub
Cancel = (UnloadMode = 1 And glbUserUploadMode = SwitchForm)
End Sub


