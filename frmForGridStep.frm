VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{AA1F4729-68B2-4E13-A27A-B298AC8511DF}#62.0#0"; "IHRCtrls.ocx"
Begin VB.Form frmForGridStep 
   Caption         =   "Update Employee Salary"
   ClientHeight    =   4845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7650
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin INFOHR_Controls.DateLookup dlpNDate 
      Height          =   285
      Left            =   2250
      TabIndex        =   2
      Tag             =   "41-Enter Salary Next Review Date"
      Top             =   1530
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.DateLookup dlpEDate 
      Height          =   285
      Left            =   2250
      TabIndex        =   1
      Tag             =   "41-Enter Salary Effective Date"
      Top             =   1170
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   503
      TextBoxWidth    =   1215
   End
   Begin INFOHR_Controls.CodeLookup clpCode 
      Height          =   285
      Index           =   1
      Left            =   2250
      TabIndex        =   0
      Tag             =   "01-Reason code "
      Top             =   810
      Width           =   4000
      _ExtentX        =   7064
      _ExtentY        =   503
      ShowUnassigned  =   1
      TABLName        =   "SDRC"
   End
   Begin Threed.SSPanel panControls 
      Align           =   2  'Align Bottom
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   4185
      Width           =   7650
      _Version        =   65536
      _ExtentX        =   13494
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
      Begin VB.CommandButton cmdCancel 
         Appearance      =   0  'Flat
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1300
         TabIndex        =   9
         Tag             =   "Update the Jobs in the following group"
         Top             =   135
         Width           =   795
      End
      Begin VB.CommandButton cmdUpdate 
         Appearance      =   0  'Flat
         Caption         =   "&OK"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Tag             =   "Update the Jobs in the following group"
         Top             =   135
         Width           =   795
      End
   End
   Begin VB.Label lblNDate 
      Caption         =   "Next Review Date"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Update Employee Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblReason 
      Caption         =   "Reason"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblEDate 
      Caption         =   "Effective Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frmForGridStep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dynSH_Job1 As New ADODB.Recordset
Dim SQLQGen
Dim fglbCOMPA#, fglbGRADE$
Dim MsgSal, IfDisplay
Dim OSalary, NSalary, OEDate, NEDate, ONDate, NNDate, EmpNo&, dblWHours#
Dim OPayp, NPayp, OJOB1, OSalCD

Private Sub cmdCancel_Click()
Dim X, Msg$
    Msg$ = "Are you sure you want to cancel updating employee salary records?"
    X = MsgBox(Msg, 36, "Confirm ")
    If X <> 6 Then Exit Sub
    
    glbGridReason = ""
    glbGridEDate = ""
    glbGridNDate = ""
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    If chkUSelect() Then
            glbGridReason = clpCode(1).Text
            glbGridEDate = dlpEDate.Text
            glbGridNDate = dlpNDate.Text
    Unload Me
    End If
End Sub

Private Function chkUSelect()
Dim Msg$, DgDef As Variant, Response%
chkUSelect = False


    If Len(clpCode(1).Text) < 1 Then
            MsgBox "Reason for Salary Change must be entered"
            clpCode(1).SetFocus
            Exit Function
    Else
      If clpCode(1).Caption = "Unassigned" Then
          MsgBox "Reason for Salary Change must be valid"
          clpCode(1).SetFocus
          Exit Function
      End If
    End If

    If Len(dlpEDate.Text) < 1 Then
        MsgBox "Effective Date must be entered"
        dlpEDate.SetFocus
        Exit Function
    Else
        If Not IsDate(dlpEDate.Text) Then
            MsgBox "Effective Date is not a valid date"
            dlpEDate.SetFocus
            Exit Function
        End If
    End If
    If Len(dlpNDate.Text) > 0 And Not IsDate(dlpNDate.Text) Then
        MsgBox "Next Review Date is not a valid date"
        dlpNDate.SetFocus
        Exit Function
    End If

chkUSelect = True


End Function
'Private Sub dlpNDate_KeyPress(KeyAscii As Integer)
'If (KeyAscii < 44 Or KeyAscii > 57) And KeyAscii <> 8 Then
'    KeyAscii = 0
'    Exit Sub
'End If
'If KeyAscii = 8 Then ActiveControl.CausesValidation = True Else ActiveControl.CausesValidation = False
'End Sub



Private Sub Form_Load()
glbOnTop = "FRMFORGRIDSTEP"
Call INI_Controls(Me)
End Sub
