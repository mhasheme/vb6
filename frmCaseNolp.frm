VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCaseNolp 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Case Information"
   ClientHeight    =   6012
   ClientLeft      =   1320
   ClientTop       =   660
   ClientWidth     =   8808
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6012
   ScaleWidth      =   8808
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4440
      Top             =   4560
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   572
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
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "frmCaseNolp.frx":0000
      Height          =   3795
      Left            =   0
      OleObjectBlob   =   "frmCaseNolp.frx":0014
      TabIndex        =   0
      Tag             =   "Department Listings"
      Top             =   0
      Width           =   8535
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   5355
      Width           =   8805
      _Version        =   65536
      _ExtentX        =   15531
      _ExtentY        =   1164
      _StockProps     =   15
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
      Begin VB.CommandButton cmdSelect 
         Appearance      =   0  'Flat
         Caption         =   "&Select"
         Height          =   375
         Left            =   60
         TabIndex        =   2
         Tag             =   "Select this Department"
         Top             =   150
         Width           =   735
      End
      Begin VB.CommandButton cmdClose 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Left            =   930
         TabIndex        =   3
         Tag             =   "Close and exit this screen"
         Top             =   150
         Width           =   735
      End
      Begin Crystal.CrystalReport vbxCrystal 
         Left            =   1860
         Top             =   150
         _ExtentX        =   593
         _ExtentY        =   593
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
End
Attribute VB_Name = "frmCaseNolp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNewRec%
'im fglbMultiSelect As Boolean
Dim rsDATA As New ADODB.Recordset ' Sam add july 02 * Remove Ado



Private Sub cmdCancel_Click()
On Error GoTo Can_Err

rsDATA.CancelUpdate
Call Display_Value



Call modSTUPD(False)    ' reset screen's attributes

cmdClose.Enabled = True
cmdClose.SetFocus

fglbNewRec% = False

Exit Sub

Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HRPROv", "Cancel")
Call RollBack '08June99

End Sub

Private Sub cmdCancel_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub cmdClose_Click()

glbCaseNum = ""
glbCaseAssociate = ""
glbCaseFiles = ""


Unload Me

End Sub

Private Sub cmdClose_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub cmdDelete_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub CmdNew_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub cmdSelect_Click()
 glbCaseFiles = ""
Dim X
If Data1.Recordset.EOF And Data1.Recordset.BOF Then
  Exit Sub
End If

If vbxTrueGrid.SelBookmarks.count <> 0 Then
    If vbxTrueGrid.SelBookmarks.count > 1000 Then
        MsgBox vbxTrueGrid.SelBookmarks.count & " codes are selected" + Chr(10) + " Please make that less than 1000 codes"
        Exit Sub
    End If
    glbCode = ""
    For X = 0 To vbxTrueGrid.SelBookmarks.count - 1
        vbxTrueGrid.Bookmark = vbxTrueGrid.SelBookmarks(X)
        glbCaseFiles = glbCaseFiles & Data1.Recordset!CaseFileNumber & ","
    Next
    glbCaseFiles = Left(glbCaseFiles, Len(glbCaseFiles) - 1)
    Unload Me
Else
    'Global glbCaseNum
            'Global glbCaseAssociate
    If Len(Data1.Recordset("CaseFileNumber")) > 0 Then
        glbCaseNum = Data1.Recordset("CaseFileNumber")
        If IsNull(Data1.Recordset("AssociationNm")) Then
            glbCaseAssociate = ""
        Else
            glbCaseAssociate = Data1.Recordset("AssociationNm")
        End If
        Unload Me
    Else
        Exit Sub
    End If
End If


End Sub

Private Sub cmdSelect_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub




Private Sub Form_Activate()
Dim xStr

    vbxTrueGrid.MultiSelect = 2
    If glbCaseFiles <> "" Then
        With Data1.Recordset
            If Not .EOF Then .MoveLast
            'Global glbCaseFiles 'Case file number Leeds & grenville - Mostafa
            'Global glbCaseNum
            'Global glbCaseAssociate
            Do Until .BOF
                If InStr(glbCaseFiles & ",", !CaseFileNumber & ",") <> 0 Then
                    xStr = Replace(xStr, !CaseFileNumber & ",", "")
                    vbxTrueGrid.SelBookmarks.Add vbxTrueGrid.Bookmark
                    DoEvents
                    If Trim(xStr) = "" Then Exit Do
                End If
                .MovePrevious
            Loop
        End With
    End If

End Sub

Private Sub Form_Load()
glbOnTop = "frmCaseNolp"
'Data1.DatabaseName = glbIHRDB
Data1.ConnectionString = GetDatabaseConStr("EFORMS")

Data1.RecordSource = "SELECT DISTINCT ltrim(rtrim(CaseFile.CaseFileNumber)) AS CaseFileNumber,  ltrim(rtrim(Association.AssociationNm)) AS AssociationNm FROM  CaseFile INNER JOIN Association ON Association.AssociationID = CaseFile.AssociationID INNER JOIN Enrollment ON CaseFile.CaseFileID = Enrollment.CaseFileID INNER JOIN EnrollmentStaff ON Enrollment.EnrollmentID = EnrollmentStaff.EnrollmentID INNER JOIN MPI ON EnrollmentStaff.StaffID = MPI.MPIID WHERE (GETDATE() BETWEEN EnrollmentStaff.StaffStart AND ISNULL(EnrollmentStaff.StaffEnd, GETDATE() + 1)) ORDER BY CaseFileNumber"
Data1.Refresh
Screen.MousePointer = HOURGLASS
Me.vbxTrueGrid.Refresh
Screen.MousePointer = DEFAULT

Call modSTUPD(False)
'Call setCaption(lblTitle(0))
'Call setCaption(lblTitle(1))
Call setCaption(Me)

Call setCaption(Me.vbxTrueGrid.Columns(0))
Call setCaption(Me.vbxTrueGrid.Columns(1))
'Call setCaption(Me.vbxTrueGrid.Columns(2))

Call INI_Controls(Me) '
End Sub

Private Sub Form_LostFocus()

MDIMain.panHelp(0).Caption = " "
MDIMain.panHelp(1).Caption = " "
MDIMain.panHelp(2).Caption = " "
MDIMain.panHelp(3).Caption = " "

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


vbxTrueGrid.Enabled = FT

cmdClose.Enabled = FT       '
cmdSelect.Enabled = FT      '

End Sub


Private Sub Form_Unload(Cancel As Integer)
Set frmCaseNolp = Nothing  'carmen may 2000
End Sub

Private Sub txtFindDesc_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtFindKey_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtFindKey_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub txtName_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub


Private Sub txtName_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(Chr$(KeyAscii)) 'Frank 5/4/2000 Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtNumber_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub



Private Sub txtNumber_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub vbxTrueGrid_DblClick()

If cmdSelect.Enabled Then
    If Not Me.vbxTrueGrid.EditActive Then
        glbCaseNum = Data1.Recordset("CaseFileNumber")
        glbCaseAssociate = Data1.Recordset("AssociationNm")
        Unload Me
    Else
        MsgBox "Save/cancel changes first"
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
        
        SQLQ = "SELECT DISTINCT ltrim(rtrim(CaseFile.CaseFileNumber)) AS CaseFileNumber,  ltrim(rtrim(Association.AssociationNm)) AS AssociationNm FROM  CaseFile INNER JOIN Association ON Association.AssociationID = CaseFile.AssociationID INNER JOIN Enrollment ON CaseFile.CaseFileID = Enrollment.CaseFileID INNER JOIN EnrollmentStaff ON Enrollment.EnrollmentID = EnrollmentStaff.EnrollmentID INNER JOIN MPI ON EnrollmentStaff.StaffID = MPI.MPIID WHERE (GETDATE() BETWEEN EnrollmentStaff.StaffStart AND ISNULL(EnrollmentStaff.StaffEnd, GETDATE() + 1)) ORDER BY CaseFileNumber"
        SQLQ = SQLQ & " ORDER BY " & vbxTrueGrid.Columns(ColIndex).DataField & " " & vbxTrueGrid.Tag
        
    
        Data1.RecordSource = SQLQ
        Data1.Refresh
'        Set FRS = Data1.Recordset.Clone
'        vbxTrueGrid.FetchRowStyle = True
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
''' Sam add July 2002 * Remove ADO
Private Sub Display_Value()
    Dim SQLQ
    If Data1.Recordset.EOF Or Data1.Recordset.BOF Then
        Call Set_Control("B", Me)
        Exit Sub
    End If

    
    SQLQ = "SELECT DISTINCT ltrim(rtrim(CaseFile.CaseFileNumber)) AS CaseFileNumber,  ltrim(rtrim(Association.AssociationNm)) AS AssociationNm FROM  CaseFile INNER JOIN Association ON Association.AssociationID = CaseFile.AssociationID INNER JOIN Enrollment ON CaseFile.CaseFileID = Enrollment.CaseFileID INNER JOIN EnrollmentStaff ON Enrollment.EnrollmentID = EnrollmentStaff.EnrollmentID INNER JOIN MPI ON EnrollmentStaff.StaffID = MPI.MPIID WHERE (GETDATE() BETWEEN EnrollmentStaff.StaffStart AND ISNULL(EnrollmentStaff.StaffEnd, GETDATE() + 1)) ORDER BY CaseFileNumber"
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, GetDatabaseConStr("EFORMS"), adOpenKeyset, adLockOptimistic
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'''Sam add July 02 * Remove ADO
Call Display_Value
End Sub



