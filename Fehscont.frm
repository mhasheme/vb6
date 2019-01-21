VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "ihrctrls.ocx"
Begin VB.Form frmEHSContact 
   AutoRedraw      =   -1  'True
   Caption         =   "Contact"
   ClientHeight    =   9000
   ClientLeft      =   -135
   ClientTop       =   600
   ClientWidth     =   9900
   ForeColor       =   &H00000000&
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   9900
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar scrControl 
      Height          =   5715
      LargeChange     =   315
      Left            =   9360
      Max             =   100
      SmallChange     =   315
      TabIndex        =   36
      Top             =   2880
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.HScrollBar scrHScroll 
      Height          =   300
      LargeChange     =   25
      Left            =   0
      Max             =   50
      SmallChange     =   4
      TabIndex        =   35
      Top             =   8640
      Width           =   9615
   End
   Begin TrueOleDBGrid60.TDBGrid vbxTrueGrid 
      Bindings        =   "Fehscont.frx":0000
      Height          =   2325
      Left            =   120
      OleObjectBlob   =   "Fehscont.frx":0014
      TabIndex        =   0
      Top             =   480
      Width           =   9075
   End
   Begin VB.TextBox Updstats 
      DataField       =   "CT_LDate"
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   2520
      MaxLength       =   25
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8670
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "CT_LTime"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   4320
      MaxLength       =   25
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8670
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox Updstats 
      DataField       =   "CT_LUSER"
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   6090
      MaxLength       =   25
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   8670
      Visible         =   0   'False
      Width           =   1590
   End
   Begin Threed.SSPanel panEEDESC 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9900
      _Version        =   65536
      _ExtentX        =   17462
      _ExtentY        =   873
      _StockProps     =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
         TabIndex        =   37
         Top             =   135
         Width           =   1305
      End
      Begin VB.Label lblEENumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Employee #"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   160
         Width           =   1005
      End
      Begin VB.Label lblEENum 
         AutoSize        =   -1  'True
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
         TabIndex        =   6
         Top             =   135
         Width           =   1245
      End
      Begin VB.Label lblEEName 
         AutoSize        =   -1  'True
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
         TabIndex        =   5
         Top             =   135
         Width           =   720
      End
   End
   Begin Crystal.CrystalReport vbxCrystal 
      Left            =   7920
      Top             =   10080
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
   Begin MSAdodcLib.Adodc Data3 
      Height          =   330
      Left            =   5760
      Top             =   10080
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   582
      ConnectMode     =   3
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
      Caption         =   "Ado3"
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
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   5640
      Top             =   10440
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   3
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
      Caption         =   "Ado1"
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
   Begin VB.Frame scrFrame 
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   9135
      Begin VB.TextBox txtAccLevel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7680
         MaxLength       =   8
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox comAccLevel 
         Height          =   315
         Left            =   6600
         TabIndex        =   15
         Tag             =   "00-Level of Accommodation"
         Top             =   1275
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox comShift 
         Height          =   315
         Left            =   2280
         TabIndex        =   34
         Tag             =   "01-Incident Number"
         Top             =   0
         Width           =   1575
      End
      Begin VB.TextBox txtShift 
         Appearance      =   0  'Flat
         DataField       =   "CT_Case"
         Height          =   285
         Left            =   3840
         MaxLength       =   8
         TabIndex        =   33
         Tag             =   "11- incident Number"
         Top             =   15
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         DataField       =   "CT_Comments"
         Height          =   1005
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Tag             =   "00-Comments"
         Top             =   4245
         Width           =   6675
      End
      Begin VB.TextBox txtCommentSuit 
         Appearance      =   0  'Flat
         DataField       =   "CT_Comment_SUIT"
         Height          =   1005
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Tag             =   "00-Suitable Work Assignment"
         Top             =   1995
         Width           =   6675
      End
      Begin VB.TextBox txtCommentRest 
         Appearance      =   0  'Flat
         DataField       =   "CT_Comment_REST"
         Height          =   1005
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Tag             =   "00-Restrictions/Capabilities"
         Top             =   3135
         Width           =   6675
      End
      Begin VB.CheckBox chkNurse 
         DataField       =   "CT_NURSE_CHK"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   1635
         Width           =   375
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "CT_FlDate"
         Height          =   285
         Index           =   1
         Left            =   1965
         TabIndex        =   14
         Tag             =   "40-Date of Follow Up"
         Top             =   1275
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.DateLookup dlpDate 
         DataField       =   "CT_CDate"
         Height          =   285
         Index           =   0
         Left            =   1960
         TabIndex        =   16
         Tag             =   "41-Date  occurred"
         Top             =   855
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   503
         TextBoxWidth    =   1215
      End
      Begin INFOHR_Controls.CodeLookup clpCode 
         DataField       =   "CT_Code"
         Height          =   285
         Index           =   0
         Left            =   1960
         TabIndex        =   17
         Tag             =   "01-Contact Type Code"
         Top             =   435
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   503
         ShowUnassigned  =   1
         TABLName        =   "ECCT"
      End
      Begin MSMask.MaskEdBox medHours 
         DataField       =   "CT_TIMESPENT"
         Height          =   285
         Left            =   6600
         TabIndex        =   19
         Tag             =   "00-Time Spent "
         Top             =   1635
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   15
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
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Level of Accommodation"
         Height          =   255
         Index           =   9
         Left            =   4560
         TabIndex        =   38
         Top             =   1305
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Image imgEmail 
         Height          =   315
         Left            =   0
         Picture         =   "Fehscont.frx":3F5C
         Stretch         =   -1  'True
         Top             =   450
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Date"
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
         Left            =   360
         TabIndex        =   32
         Top             =   915
         Width           =   1875
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Incident Number"
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
         Left            =   360
         TabIndex        =   31
         Top             =   75
         Width           =   1545
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Type"
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
         Index           =   4
         Left            =   360
         TabIndex        =   30
         Top             =   495
         Width           =   1695
      End
      Begin VB.Label lblTitle 
         Caption         =   "Comments"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   29
         Top             =   4275
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Caption         =   "Follow Up Date"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   1305
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Caption         =   "Suitable Work Assignment"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   27
         Top             =   2025
         Width           =   1905
      End
      Begin VB.Label lblTitle 
         Caption         =   "Restrictions/Capabilities"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   26
         Top             =   3165
         Width           =   1845
      End
      Begin VB.Label lblTitle 
         Caption         =   "Nurse"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   25
         Top             =   1635
         Width           =   1305
      End
      Begin VB.Label lblTitle 
         Alignment       =   1  'Right Justify
         Caption         =   "Time Spent"
         Height          =   255
         Index           =   8
         Left            =   5490
         TabIndex        =   24
         Top             =   1650
         Width           =   945
      End
      Begin VB.Label lblUpdateDate 
         Caption         =   "Updated Date"
         Height          =   255
         Left            =   5610
         TabIndex        =   23
         Top             =   5355
         Width           =   1095
      End
      Begin VB.Label lblUpdDateDesc 
         Height          =   255
         Left            =   6690
         TabIndex        =   22
         Top             =   5355
         Width           =   1935
      End
      Begin VB.Label lblUpdateBy 
         Caption         =   "Updated By"
         Height          =   255
         Left            =   2250
         TabIndex        =   21
         Top             =   5355
         Width           =   975
      End
      Begin VB.Label lblUserDesc 
         Height          =   255
         Left            =   3210
         TabIndex        =   20
         Top             =   5355
         Width           =   2415
      End
   End
   Begin VB.Label lblEEID 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CT_Empnbr"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   1470
      TabIndex        =   8
      Top             =   8790
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblCNum 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label20"
      DataField       =   "CT_CompNo"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   570
      TabIndex        =   9
      Top             =   8760
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmEHSContact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fglbNew
Dim fglHredsem
Dim rsDATA As New ADODB.Recordset ' Sam add July 2002 * Remove Binding Control

Function chkHSContact()

Dim SQLQ As String, Msg As String, dd#

chkHSContact = False

On Error GoTo chkHSContact_Err
If Len(dlpDate(0).Text) >= 1 Then
    If Not IsDate(dlpDate(0).Text) Then
        MsgBox "Contact Date Not Valid."
        dlpDate(0).SetFocus
        Exit Function
    End If
Else
    MsgBox "Contact Date is required."
    dlpDate(0).SetFocus
    Exit Function
End If
If Len(dlpDate(1).Text) > 0 Then
    If Not IsDate(dlpDate(1).Text) Then
        MsgBox "Follow up Date Not Valid."
        dlpDate(1).SetFocus
        Exit Function
    End If
    If DateDiff("d", CVDate(dlpDate(0).Text), CVDate(dlpDate(1).Text)) < 0 Then
        Msg$ = "Follow up date precede Contact date"
        dlpDate(1).SetFocus
        MsgBox Msg$
        Exit Function
    End If
End If
Dim tTime As Variant
Dim Part1$, Part2$

'~~

If Len(txtShift) < 1 Then
    MsgBox "Incident Number is a required field"
    comShift.SetFocus
    Exit Function
End If
If Not IfIncidentNo(Val(txtShift)) Then
    MsgBox "Incident Number Not Valid"
    comShift.SetFocus
    Exit Function
End If

If Len(clpCode(0).Text) < 1 Then
    MsgBox "Contact Code is a required field"
    clpCode(0).SetFocus
    Exit Function
End If

If clpCode(0).Caption = "Unassigned" Then
    MsgBox "Contact code must be valid"
    clpCode(0).SetFocus
    Exit Function
End If

If chkNurse.Value Then
    If Len(medHours.Text) = 0 Then
        MsgBox "Time Spent is required when Nurse is checked!"
        medHours.SetFocus
        Exit Function
    End If
End If

chkHSContact = True

Exit Function

chkHSContact_Err:

glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "chkHSInc", "HR_OHS_CONTACT", "edit/Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If


End Function

'Private Sub cmdCAction_Click()
'frmEHSCorrective.Show
'Unload Me
'End Sub

Sub cmdCancel_Click()
Dim x
On Error GoTo Can_Err

fglbNew = False
Call Display_Value
'Call ST_UPD_MODE(True)  ' reset screen's attributes
'Call SET_UP_MODE
Me.vbxTrueGrid.SetFocus
Exit Sub
Can_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err
Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Can Error", "HR_OHS_CONTACT", "Cancel")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

'Sub cmdCancel_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdClose_Click()
Call NextForm
Unload Me
If glbOnTop = "FRMEHSCONTACT" Then glbOnTop = ""

End Sub

'Sub cmdClose_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdDelete_Click()
Dim a As Integer, Msg As String, INo&, x

fglHredsem = dlpDate(1).Text

If Not gSec_Upd_HSContacts Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
    MsgBox "Nothing to Delete"
    Exit Sub
End If


On Error GoTo Del_Err


Msg = "Are You Sure You Want To Delete "
Msg = Msg & Chr(10) & "This Record?  "

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

If Not glbtermopen Then
    If Not updFollow("D") Then
        Exit Sub
    End If
End If

Me.vbxTrueGrid.SetFocus
fglbNew = False
'Call ST_UPD_MODE(True)
Call SET_UP_MODE
Exit Sub

Del_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdDel", "HR_OHS_CONTACT", "Delete")
Screen.MousePointer = DEFAULT
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If
End Sub

Sub cmdNew_Click()
Dim SQLQ As String

If Not gSec_Upd_HSContacts Then
    MsgBox "You Do Not Have Authority For This Transacaction"
    Exit Sub
End If

fglbNew = True

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

On Error GoTo AddN_Err

fglbNew = True

'data1.Recordset.AddNew
''' Sam add July 2002 * Remove Binding Control

Call Set_Control("B", Me)


fglHredsem = ""
If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID

lblCNum.Caption = "001"
comShift.SetFocus

Exit Sub

AddN_Err:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdAdd", "HR_OHS_CONTACT", "Add")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Sub CmdNew_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdOK_Click()
Dim x
On Error GoTo Add_Err

If Not chkHSContact() Then Exit Sub
rsDATA.Requery

If fglbNew Then
    rsDATA.AddNew
    rsDATA("CT_CODE_TABL") = "ECCT"
End If

If Not glbtermopen Then
    'Ticket #22682: Release 8.0 - Set older Follow Up records as Completed first if uncompleted
    'follow up records are found for Salary, before adding a new follow up record.
    If fglbNew Then
        glbFollowUpList = "HSFU"
        If Older_FollowUp_Records_Found(glbFollowUpList) Then
            frmFollowUpList.Show 1
        End If
    End If

    If Not updFollow("U") Then
        Exit Sub
    End If
    
    Call UpdUStats(Me)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001.BeginTrans
    rsDATA.Update
    gdbAdoIhr001.CommitTrans
Else
    rsDATA!TERM_SEQ = glbTERM_Seq
    Call UpdUStats(Me)
    Call Set_Control("U", Me, rsDATA)
    gdbAdoIhr001X.BeginTrans
    rsDATA.Update
    gdbAdoIhr001X.CommitTrans

End If
Data1.Refresh
fglbNew = False

'Call ST_UPD_MODE(True)
Call SET_UP_MODE

If NextFormIF("Contact") Then
    Call cmdNew_Click
End If

Exit Sub

Add_Err:
If Err = 3022 Then
    'Data1.UpdateControls  ' no dups
    MsgBox "Duplicate record existed - not entered"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Sub
End If
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "cmdUpdate", "HR_OHS_CONTACT", "Update")
If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

End Sub

'Sub cmdOK_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

Sub cmdPrint_Click()
Dim RHeading As String

RHeading = lblEEName & "'s Contact"
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

RHeading = lblEEName & "'s Contact"
Me.vbxCrystal.WindowTitle = RHeading & " Report"
Me.vbxCrystal.BoundReportHeading = RHeading
Me.vbxCrystal.Destination = 0
Me.vbxCrystal.Action = 1
End Sub

Private Sub comAccLevel_Click()
txtAccLevel.Text = comAccLevel.Text
End Sub

Private Sub comAccLevel_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

'Sub cmdPrint_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdTCause_Click()
'frmEHSCause.Show
'Unload Me
'End Sub

'Private Sub cmdTCause_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Sub cmdWCBMed_Click()
'frmEHSWCB.Show
'Unload Me
'End Sub
'Private Sub cmdWCBMed_GotFocus()
'Call SetPanHelp(ActiveControl)
'End Sub

'Private Sub cmdWSIB_Click()
'frmEHSWCBC.Show
'Unload Me
'End Sub

Sub comShift_Change()
'txtShift = comShift  'JDY
End Sub

Sub comShift_Click()
'txtShift = comShift      'JDY
End Sub

Sub comShift_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub comShift_LostFocus()
txtShift = comShift      'JDY
End Sub

Sub Data1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)

glbFrmCaption$ = Me.Caption
glbErrNum& = ErrorNumber

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "DATA1.error", "HR_OHS_CONTACT", "SELECT")

End Sub

Function EERetrieve()

Dim SQLQ As String

EERetrieve = False

Screen.MousePointer = HOURGLASS
On Error GoTo EERError


If glbtermopen Then         'Lucy July 5, 2000
    SQLQ = "Select * from Term_OHS_CONTACT"
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    'Ticket #22678 - Granite Club - Sort by Contact Date
    If glbCompSerial = "S/N - 2241W" Then
        SQLQ = SQLQ & " ORDER BY CT_CDate DESC"
    Else
        SQLQ = SQLQ & " ORDER BY CT_Case DESC"
    End If
Else
    SQLQ = "Select * from HR_OHS_CONTACT "
    SQLQ = SQLQ & " where CT_Empnbr = " & glbLEE_ID
    'Ticket #22678 - Granite Club - Sort by Contact Date
    If glbCompSerial = "S/N - 2241W" Then
        SQLQ = SQLQ & " ORDER BY CT_CDate DESC"
    Else
        SQLQ = SQLQ & " ORDER BY CT_Case DESC"
    End If
End If


Data1.RecordSource = SQLQ
Data1.Refresh

If glbtermopen Then     'Lucy July 5, 2000
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from Term_HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
Else
    SQLQ = "SELECT EC_EMPNBR, EC_CASE, EC_OCCDATE from HR_OCC_HEALTH_SAFETY "
    SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " ORDER BY EC_CASE DESC"
End If

Data3.RecordSource = SQLQ
Data3.Refresh



EERetrieve = True
Screen.MousePointer = DEFAULT

Exit Function


EERError:
glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "OCH Retrieve", "HR_OHS_CONTACT", "SELECT")

If gintRollBack% = False Then
    Resume Next
Else
    Unload Me
End If

Exit Function
End Function

Sub Form_Activate()
Call SET_UP_MODE
glbOnTop = "FRMEHSCONTACT"
End Sub

Sub Form_GotFocus()
glbOnTop = "FRMEHSCONTACT"
End Sub

Sub Form_Load()

Dim Answer, DefVal, Msg, Title  ' Declare variables.
Dim RFound As Integer ' records found
Dim x%
Dim SQLQ1
glbOnTop = "FRMEHSCONTACT"

If glbtermopen Then         'Lucy July 5, 2000
    Data1.ConnectionString = glbAdoIHRAUDIT
    Data3.ConnectionString = glbAdoIHRAUDIT
Else
    Data1.ConnectionString = glbAdoIHRDB
    Data3.ConnectionString = glbAdoIHRDB
End If


If glbLinHS Then 'Ticket #12401
    glbLinEmpNo = glbLEE_ID
    If Not glbtermopen Then
        If Len(glbDiv) = 0 Then Call Get_Div(False) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    Else
        If Len(glbDiv) = 0 Then Call Get_Div(False) 'frmDIVISIONS.Show 1
        If Len(glbDiv) = 0 Then Unload Me: Exit Sub
    End If
    glbLinHSDivNo = Val("999999" & glbDiv)
    glbLEE_ID = glbLinHSDivNo
    glbLEE_SName = glbDivDesc
Else
    If glbLinamar Then
        If glbLEE_ID <> 0 Then
            If Left(Trim(str(glbLEE_ID)), 6) = "999999" Then
                glbLEE_ID = 0
            End If
        End If
    End If
    If Not glbtermopen Then
        If glbLEE_ID = 0 Then frmEEFIND.Show 1
        If glbLEE_ID = 0 Then Unload Me: Exit Sub
    Else
        If glbTERM_ID = 0 Then frmTERMEMPL.Show 1
        If glbTERM_ID = 0 Then Unload Me: Exit Sub
    End If
End If

If glbLinamar Then 'Ticket #15172
    lblTitle(3).Caption = "Modified Job Duties"
    txtCommentSuit.Tag = "00-Modified Job Duties"
    lblTitle(9).Visible = True
    txtAccLevel.DataField = "CT_LEVEL"
    comAccLevel.AddItem "1"
    comAccLevel.AddItem "2"
    comAccLevel.AddItem "3"
    comAccLevel.AddItem "4"
    comAccLevel.AddItem "5"
    comAccLevel.AddItem ""
    comAccLevel.Visible = True
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

comShift.Clear
Do Until Data3.Recordset.EOF                  'JDY
  comShift.AddItem Data3.Recordset("EC_CASE") 'JDY
  Data3.Recordset.MoveNext                    'JDY
Loop

Me.vbxTrueGrid.SetFocus
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Contact Data - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    lblEENumber.Caption = lStr("Division")
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then  ' dont do on add new until in
        Me.Caption = "Contact Data - " & Left$(glbLEE_SName, 8)
        Me.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    lblEENum.Caption = ShowEmpnbr(lblEEID)
End If

Call ST_UPD_MODE(False)

Call Display_Value

If Not gSec_Upd_HSContacts Then
'    cmdModify.Enabled = False
'    cmdNew.Enabled = False
'    cmdDelete.Enabled = False
End If

Call INI_Controls(Me)

'Ticket# 13413 for Bird Packaging Limited
If glbCompSerial = "S/N - 2387W" Then
    imgEmail.Visible = True
End If

MDIMain.lstPanel.Visible = False
MDIMain.lstView.Visible = False
Screen.MousePointer = DEFAULT
End Sub

Sub Form_LostFocus()
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

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized And MDIMain.WindowState <> vbMinimized Then
    'Vertical scroll bar
    If Me.Height >= 9350 Then
        scrControl.Value = 0
        scrFrame.Top = 3200
        scrControl.Visible = False
    Else
        scrControl.Visible = True
        If Me.Height < 7000 Then
            scrControl.Max = 5000
        Else
            scrControl.Max = 3000
        End If
        scrControl.Left = Me.Width - scrControl.Width - 120
        If Me.Height - scrControl.Top - 780 > 0 Then
            scrControl.Height = Me.Height - scrControl.Top - 780
        End If
    End If
    
    'Horizontal Scroll
    scrHScroll.Width = Me.Width - 120
    'scrFrame.Height = Me.ScaleHeight - (scrHScroll.Height + 200)
    If Me.Width >= 9750 Then
        scrHScroll.Value = 0
        scrHScroll.Visible = False
    Else
        scrHScroll.Visible = True
        If Me.Width < 7000 Then
            scrHScroll.Max = 100
        Else
            scrHScroll.Max = 30
        End If
        scrHScroll.Top = Me.Height - 800
        scrHScroll.Width = Me.Width - 120
    End If
End If
End Sub

Sub Form_Unload(Cancel As Integer)

MDIMain.panHelp(0).Caption = "Select function from the menu."
Set frmEHSContact = Nothing  'carmen may 00
Call NextForm
End Sub

Function IfIncidentNo(InciNo As Double)
  IfIncidentNo = False
  If Data3.Recordset.BOF And Data3.Recordset.EOF Then
     Exit Function
  End If
  Data3.Recordset.MoveFirst
  Data3.Recordset.Find "EC_Case=" & InciNo
  If Data3.Recordset.EOF Then Exit Function
  IfIncidentNo = True

End Function

Sub ST_UPD_MODE(YN)
Dim TF As Integer, FT As Integer

If YN Then
    TF = True
    FT = False
Else
    TF = False
    FT = True
End If

glbOHSEdit% = TF

dlpDate(0).Enabled = TF
dlpDate(1).Enabled = TF
txtComments.Enabled = TF
comShift.Enabled = TF
clpCode(0).Enabled = TF

'cmdOK.Enabled = TF
'cmdCancel.Enabled = TF

'cmdClose.Enabled = FT
'cmdModify.Enabled = FT
'cmdNew.Enabled = FT
'cmdDelete.Enabled = FT
'cmdPrint.Enabled = FT
'vbxTrueGrid.Enabled = FT

'cmdWCBMed.Enabled = FT
'cmdIncident.Enabled = FT
'cmdTCause.Enabled = FT
'cmdCAction.Enabled = FT
'cmdInjLoc.Enabled = FT
'cmdWSIB.Enabled = FT

If Data1.Recordset.BOF And Data1.Recordset.EOF Then
'    cmdModify.Enabled = False
'    cmdDelete.Enabled = False
Else
    'cmdModify.Enabled = True
End If
End Sub

Private Sub medHours_GotFocus()
    Call SetPanHelp(ActiveControl)
End Sub

Private Sub scrControl_Change()
scrFrame.Top = 3120 - scrControl.Value
End Sub

Private Sub scrHScroll_Change()
scrFrame.Left = 0 - (scrHScroll.Value / 100) * ScaleWidth
End Sub

Private Sub txtAccLevel_Change()
comAccLevel.Text = txtAccLevel.Text
End Sub

Private Sub txtCommentRest_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Private Sub txtComments_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub
'Private Sub txtDate_Change(Index As Integer)
'Call Date_Change(ActiveControl)
'End Sub
'Private Sub txtDate_DblClick(Index As Integer)
'Call ShowDate(Me, Me.ActiveControl)
'End Sub
'Sub txtDate_GotFocus(Index As Integer)
'Call SetPanHelp(ActiveControl)
'End Sub
'Private Sub txtDate_KeyPress(Index As Integer, KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub

Private Sub txtCommentSuit_GotFocus()
Call SetPanHelp(ActiveControl)
End Sub

Sub txtShift_Change()
 
  If Not (Val(txtShift) = 0) Then
    comShift = txtShift
  Else
    comShift = ""
  End If

End Sub

Private Sub Updstats_Change(Index As Integer)
    If Index = 0 Then
        'If IsDate(Updstats(Index).Text) Then
        lblUpdDateDesc.Caption = Updstats(Index).Text
        'End If
    End If
    If Index = 2 Then
        lblUserDesc.Caption = GetUserDesc(Updstats(Index))
    End If
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
            SQLQ = "Select * from Term_OHS_CONTACT"
            SQLQ = SQLQ & " WHERE TERM_SEQ = " & glbTERM_Seq
        Else
            SQLQ = "Select * from HR_OHS_CONTACT "
            SQLQ = SQLQ & " where CT_Empnbr = " & glbLEE_ID
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
'    If cmdOK.Enabled Then
'        cmdOK.SetFocus
'    Else
'        cmdModify.SetFocus
'    End If
End If

End Sub

Private Function updFollow(xType)   'Laura on 11/2/97
Dim newline As String
Dim SQLQ As String
Dim Msg As String
Dim rsTB As New ADODB.Recordset
Dim dynHRAT As New ADODB.Recordset
Dim rsTT As New ADODB.Recordset
Dim Edit1 As Integer
'Don't need a message for follow up - Jerry asked for v7.6

newline = Chr$(13) & Chr$(10)
updFollow = False

On Error GoTo CrFollow_Err

If fglHredsem <> "" Then    'DATE Renewal IS NOW MANDATORY
    SQLQ = "SELECT * FROM HR_FOLLOW_UP WHERE EF_EMPNBR = " & glbLEE_ID
    SQLQ = SQLQ & " AND EF_FREAS = 'HSFU'"
    SQLQ = SQLQ & " AND EF_FDATE = " & Date_SQL(fglHredsem)

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
    If fglbNew And dlpDate(1).Text <> "" Then
        'Create the Code if not already existing
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='HSFU'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "HSFU"
            rsTT("TB_DESC") = "Health & Safety " & lStr("Follow-ups")
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        Set rsTT = Nothing

        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "HSFU", "Health & Safety " & lStr("Follow-ups"))
        
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "HSFU"
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
        'tkt#10995
        'Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
    If fglbNew = False And Edit1 = False And dlpDate(1).Text <> "" Then
        'Create the Code if not already existing
        rsTT.Open "SELECT * FROM HRTABL WHERE TB_NAME='FURE' AND TB_KEY='HSFU'", gdbAdoIhr001, adOpenStatic, adLockOptimistic
        If rsTT.EOF Then
            rsTT.AddNew
            rsTT("TB_COMPNO") = "001"
            rsTT("TB_NAME") = "FURE"
            rsTT("TB_KEY") = "HSFU"
            rsTT("TB_DESC") = "Health & Safety " & lStr("Follow-ups")
            rsTT("TB_LUSER") = glbUserID
            rsTT("TB_LDATE") = Date
            rsTT("TB_LTIME") = Time$
            rsTT.Update
        End If
        rsTT.Close
        Set rsTT = Nothing

        'Release 8.0 - Grant permission to this Follow Up for this user as well so the user can see the
        'follow up record
        Call Grant_FollowUpCode_Security(glbUserID, "HSFU", "Health & Safety " & lStr("Follow-ups"))
        
        rsTB.AddNew
        rsTB("EF_COMPNO") = "001"
        rsTB("EF_EMPNBR") = glbLEE_ID
        rsTB("EF_FDATE") = CVDate(dlpDate(1).Text)
        rsTB("EF_FREAS_TABL") = "FURE"
        'Ticket #24257 - Do not update Admin By for them only
        If glbCompSerial <> "S/N - 2262W" Then
            rsTB("EF_ADMINBY_TABL") = "EDAB"
            rsTB("EF_ADMINBY") = GetEmpData(glbLEE_ID, "ED_ADMINBY", Null)
        End If
        rsTB("EF_FREAS") = "HSFU"
        rsTB("EF_COMMENTS") = ""
        rsTB("EF_LDATE") = Date
        rsTB("EF_LTIME") = Time$
        rsTB("EF_LUSER") = glbUserID
        rsTB.Update
        rsTB.Close
        updFollow = True
         'tkt#10995
       ' Msg = "A Follow Up Record was created!"
        'MsgBox Msg
        Exit Function
    End If
  
    If fglbNew = False And Edit1 = True And dlpDate(1).Text <> "" Then  ' edited record
        'EOF?
        dynHRAT.MoveFirst
        Do Until dynHRAT.EOF
            'dynHRAT.Edit
            dynHRAT("EF_COMPNO") = "001"
            dynHRAT("EF_EMPNBR") = glbLEE_ID
            dynHRAT("EF_FDATE") = CVDate(dlpDate(1).Text)
            dynHRAT("EF_FREAS") = "HSFU"
            dynHRAT("EF_COMMENTS") = ""
            dynHRAT("EF_LDATE") = Date
            dynHRAT("EF_LTIME") = Time$
            dynHRAT("EF_LUSER") = glbUserID
            dynHRAT.Update
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        If fglHredsem <> dlpDate(1).Text Then
           'tkt#10995
            'Msg = "A Follow Up Record was updated!"
            'MsgBox Msg
        End If
        updFollow = True
        Edit1 = True
        Exit Function
    End If
    If fglbNew = False And Edit1 = True And dlpDate(1).Text = "" Then
        Do Until dynHRAT.EOF
            dynHRAT.Delete
            dynHRAT.MoveNext
        Loop
        dynHRAT.Close
        Edit1 = True
        updFollow = True
         'tkt#10995
       ' Msg = "A record has been deleted from the Follow Up table"
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
        'Msg = "A record has been deleted from the Follow Up table"
        'MsgBox Msg
        Exit Function
    Else
        updFollow = True
    End If
End If

If dlpDate(1).Text = "" Then
    updFollow = True
End If
  
Exit Function

CrFollow_Err:
If Err = 3022 Then
    MsgBox "The record is not entered or deleted!"
    Err = 0   ' i know will be reset any way - but just in case
    Resume Next
    Exit Function
End If

glbFrmCaption$ = Me.Caption
glbErrNum& = Err

Call ERR_Hndlr(glbErrNum&, glbFrmCaption$, "Follow UP", "HR_FOLLOW_UP", "UPDATE TABLE")
Resume Next

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
        Call SET_UP_MODE
        Exit Sub
    End If
    
    
If glbtermopen Then
    SQLQ = "Select * from Term_OHS_CONTACT"
    SQLQ = SQLQ & " WHERE CT_ID = " & Data1.Recordset!CT_ID
    'Ticket #22678 - Granite Club - Sort by Contact Date
    If glbCompSerial = "S/N - 2241W" Then
        SQLQ = SQLQ & " ORDER BY CT_CDate DESC"
    Else
        SQLQ = SQLQ & " ORDER BY CT_Case DESC"
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001X, adOpenKeyset, adLockOptimistic
Else
    SQLQ = "Select * from HR_OHS_CONTACT "
    SQLQ = SQLQ & " where CT_ID = " & Data1.Recordset!CT_ID
    'Ticket #22678 - Granite Club - Sort by Contact Date
    If glbCompSerial = "S/N - 2241W" Then
        SQLQ = SQLQ & " ORDER BY CT_CDate DESC"
    Else
        SQLQ = SQLQ & " ORDER BY CT_Case DESC"
    End If
    If rsDATA.State <> 0 Then: If rsDATA.EOF Then rsDATA.Close Else If rsDATA.EditMode = adEditAdd Then rsDATA.CancelUpdate: rsDATA.Close Else rsDATA.Close
    rsDATA.Open SQLQ, gdbAdoIhr001, adOpenKeyset, adLockOptimistic
End If
    
    If rsDATA.EOF Or rsDATA.BOF Then Exit Sub
    Call Set_Control("R", Me, rsDATA)
Call SET_UP_MODE

'Ticket #22682 - Bug fix. The Follow Up record should be edited when the Follow Up Date is changed on an existing record.
fglHredsem = dlpDate(1).Text

End Sub

Private Sub vbxTrueGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Call Display_Value

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
UpdateRight = gSec_Upd_HSContacts
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
If glbLinHS Then
    If Len(glbDivDesc) > 0 Then   ' dont do on add new until in
        Me.Caption = "Contacts - " & glbDivDesc
        Me.lblEEName = RTrim$(glbDivDesc)
    End If
    lblEENum.Caption = glbDiv
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = ""
    End If
Else
    If Len(glbLEE_SName) > 0 And Len(glbLEE_SName) > 0 Then
        frmEHSContact.Caption = "Contacts - " & Left$(glbLEE_SName, 5)
        frmEHSContact.lblEEName = RTrim$(glbLEE_SName) & ", " & RTrim$(glbLEE_FName)
    End If
    If glbtermopen Then lblEEID = glbTERM_ID Else lblEEID = glbLEE_ID
    'lblEEID = glbLEE_ID
    lblEENum = ShowEmpnbr(lblEEID)
    
    If glbLinamar Then  'Ticket #14775
        lblEEProdLine = glbLEE_ProdLine
    Else
        lblEEProdLine = ""
    End If
End If
End Sub

Public Sub imgEmail_Click()
    Call EmailShow
End Sub

Private Sub EmailShow()
Dim rsTMail As New ADODB.Recordset
Dim xReportByEmail As String
Dim xAssignNO
Dim SQLQ As String

On Error GoTo Email_Err
    If glbtermopen Then Exit Sub
    If gsEMAIL_SENDING Then
        xReportByEmail = ""
        SQLQ = "SELECT EC_SUPERVISOR FROM HR_OCC_HEALTH_SAFETY "
        SQLQ = SQLQ & " WHERE EC_EMPNBR = " & glbLEE_ID
        
        rsTMail.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If rsTMail.EOF Then
            Exit Sub
        End If
        If IsNull(rsTMail("EC_SUPERVISOR")) Then
            MsgBox "No Assigned To on Incident Screen"
            Exit Sub
        End If
        If Len(rsTMail("EC_SUPERVISOR")) = 0 Then
            MsgBox "No Assigned To on Incident Screen"
            Exit Sub
        End If
        xAssignNO = rsTMail("EC_SUPERVISOR")
        rsTMail.Close
        
        SQLQ = "SELECT ED_EMPNBR, ED_EMAIL FROM HREMP WHERE ED_EMPNBR = " & xAssignNO & " "
        rsTMail.Open SQLQ, gdbAdoIhr001, adOpenStatic
        If Not rsTMail.EOF Then
            If Not IsNull(rsTMail("ED_EMAIL")) Then
                If Len((rsTMail("ED_EMAIL"))) Then
                    xReportByEmail = rsTMail("ED_EMAIL")
                End If
            End If
        End If
        rsTMail.Close
        If Len(xReportByEmail) > 0 Then
            frmSendEmail.txtTo.Text = xReportByEmail
            frmSendEmail.Tag = ""
            frmSendEmail.Show 1
        Else
            MsgBox "Assigned To Email Address is blank."
        End If
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

