VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSelDate 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendar"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2730
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      MaxSelCount     =   2000
      MultiSelect     =   -1  'True
      StartOfWeek     =   61734913
      CurrentDate     =   36658
   End
End
Attribute VB_Name = "frmSelDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
If IsDate(trsDate.TextBox) Then MonthView1 = trsDate.TextBox Else MonthView1 = Now
Me.Caption = lStr(Mid(trsDate.TextBox.Tag, InStr(trsDate.TextBox.Tag, "-") + 1))

End Sub

Private Sub Form_Load()
If trsDate.TextBox.Parent.Caption = trsDate.TextBox.Container.Caption Then
    Me.Left = trsDate.Form.Left + trsDate.TextBox.Left + 50
    Me.Top = trsDate.Form.Top + trsDate.TextBox.Top + 1200
Else
    Me.Left = trsDate.Form.Left + trsDate.TextBox.Container.Left + trsDate.TextBox.Left + 50
    Me.Top = trsDate.Form.Top + trsDate.TextBox.Container.Top + trsDate.TextBox.Top + 1200
End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    trsDate.TextBox = DateClicked
    If MonthView1.SelStart <> MonthView1.SelEnd Then
        trsDate.Form.txtToDate = MonthView1.SelEnd
    End If
    Unload Me
End Sub


Private Sub MonthView1_LostFocus()
    Unload Me
End Sub



